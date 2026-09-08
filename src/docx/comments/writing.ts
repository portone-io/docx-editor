/**
 * Plans changed comment package parts while preserving untouched XML.
 */

import type { Node as PMNode } from "prosemirror-model";
import { elementXml, type XmlAttr } from "../../ooxml/element";
import { NAMESPACES, wName, xmlnsDecl } from "../../ooxml/names";
import {
  ensureRootDeclarations,
  partRootProblem,
  type RootDeclarations,
  splicePart,
} from "../../ooxml/partSplice";
import { encodeUtf8 } from "../../ooxml/xml";
import type { PartPlanContext, PartPlanner } from "../partPlan";
import { directoryOf } from "../relationships";
import type { SessionStore } from "../session";
import {
  arrivedEntries,
  renderCommentBody,
  renderCommentExtension,
  withThreadKey,
} from "./grammar";
import {
  type CommentReferenceData,
  type CommentReplyData,
  commentReferencesIn,
} from "./model";
import { commentsExtendedPart, commentsPart, peoplePart } from "./parts";
import { planPeoplePart } from "./people";
import type { ImportedComments } from "./reading";

/**
 * Whether the entry this comment arrived as has to gain the key its thread state hangs off.
 *
 * A comment settled or replied to for the first time is written into the extended part under a
 * key, and an entry that arrived without one has none for that to name.
 */
function needsThreadKey(
  id: string,
  comment: CommentReferenceData,
  session: SessionStore
): boolean {
  return (
    comment.imported &&
    carriesThreadMetadata(comment) &&
    (session.comments.byId.get(id)?.paraId ?? null) === null
  );
}

/**
 * Whether the Comments part has to be rewritten: a comment or reply came, went, or changed, or an
 * entry has to gain the key its thread state hangs off. The export invariants ask this as well, so
 * a part the writer would leave alone is one they do not read.
 */
export function commentsChanged(doc: PMNode, session: SessionStore): boolean {
  const current = commentReferencesIn(doc);
  if (current.size !== session.commentReferenceIds.size) return true;
  const currentBodies = currentCommentBodies(current);
  const originalBodies = originalThreadIds(
    session.comments,
    session.commentReferenceIds
  );
  if (currentBodies.size !== originalBodies.size) return true;
  for (const id of originalBodies) {
    if (!currentBodies.has(id)) return true;
  }
  for (const id of session.commentReferenceIds) {
    const comment = current.get(id);
    if (!comment?.imported) return true;
  }
  for (const comment of current.values()) {
    if (comment.replies.some((reply) => !reply.imported)) return true;
  }
  for (const [id, comment] of current) {
    if (needsThreadKey(id, comment, session)) return true;
  }
  return false;
}

function extensionsXml(
  references: ReadonlyMap<string, CommentReferenceData>,
  comments: ImportedComments,
  originallyReferenced: ReadonlySet<string>
): string {
  const current = currentCommentBodies(references);
  const originalThreads = originalThreadIds(comments, originallyReferenced);
  const pieces: string[] = [];
  const written = new Set<string>();
  const idByParaId = new Map(
    comments.ordered.flatMap((comment) =>
      comment.paraId === null ? [] : [[comment.paraId, comment.id] as const]
    )
  );

  for (const original of comments.extendedOrdered) {
    const id = idByParaId.get(original.paraId);
    const item = id === undefined ? undefined : current.get(id);
    if (id !== undefined && item && !written.has(id)) {
      pieces.push(renderCommentExtension(item));
      written.add(id);
    } else if (id === undefined || !originalThreads.has(id)) {
      pieces.push(original.xml);
    }
  }
  for (const [id, comment] of current) {
    // A comment with no thread state carries no key for an entry here to name
    if (!written.has(id) && carriesThreadMetadata(comment)) {
      pieces.push(renderCommentExtension(comment));
    }
  }

  if (comments.extendedXml === null) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w15:commentsEx ${xmlnsDecl("w15")}>${pieces.join("")}</w15:commentsEx>`
    );
  }
  return ensureRootDeclarations(
    splicePart(comments.extendedXml, {
      root: EXTENSIONS_ROOT,
      replaceChildren: pieces.join(""),
    }),
    { namespaces: { w15: NAMESPACES.w15 } }
  );
}

/** The root element each comment part is rewritten around */
const COMMENTS_ROOT = "comments";
const EXTENSIONS_ROOT = "commentsEx";

/**
 * Why the Comments part cannot be rewritten around its root, or null when it can.
 *
 * The same question stands in the export invariant list, so a refusal the writer raises is one a
 * caller could have read there first, in the same words.
 */
export function commentsRootProblem(xml: string): string | null {
  return partRootProblem(xml, COMMENTS_ROOT);
}

/** Why the extended part cannot be rewritten around its root, or null when it can */
export function extensionsRootProblem(xml: string): string | null {
  return partRootProblem(xml, EXTENSIONS_ROOT);
}

/** Whether the extended part has to be rewritten: thread state arrived, changed, or went with a deleted comment */
export function extensionsChanged(doc: PMNode, session: SessionStore): boolean {
  const current = commentReferencesIn(doc);
  for (const [id, comment] of current) {
    if (!comment.threadImported) return true;
    if (!session.commentReferenceIds.has(id)) continue;
  }
  for (const id of session.commentReferenceIds) {
    if (current.has(id)) continue;
    const original = session.comments.byId.get(id);
    if (
      original?.extensionXml != null ||
      (session.comments.repliesByParentId.get(id)?.length ?? 0) > 0
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Whether the comment has thread state to write down: it arrived with some, it has been settled
 * or replied to since, or it is itself a reply. A comment with none needs no entry in the extended
 * part, where absent reads as an open thread standing on its own.
 */
function carriesThreadMetadata(
  comment: CommentReferenceData | CommentReplyData
): boolean {
  return (
    comment.extensionXml !== null ||
    ("threadImported" in comment && !comment.threadImported) ||
    "parentParaId" in comment
  );
}

/**
 * The thread key belongs on the entry where the comment has thread state to hang off it, and
 * where the entry arrived carrying one: a key already written is what its state is keyed by
 * elsewhere, so a rewrite of what the comment says keeps it.
 */
function keyedEntry(
  comment: CommentReferenceData | CommentReplyData,
  arrivedKeyed: ReadonlySet<string>
): boolean {
  return carriesThreadMetadata(comment) || arrivedKeyed.has(comment.id);
}

function renderedComment(
  comment: CommentReferenceData | CommentReplyData,
  arrivedKeyed: ReadonlySet<string>,
  arrived: ReadonlyMap<string, Element>
): string {
  if (comment.imported && comment.commentXml !== null) {
    return carriesThreadMetadata(comment)
      ? withThreadKey(
          comment.commentXml,
          comment.paraId,
          arrived.get(comment.id) ?? null
        )
      : comment.commentXml;
  }
  const attrs: readonly (XmlAttr | null)[] = [
    [wName("id"), comment.id],
    comment.author === null ? null : [wName("author"), comment.author],
    comment.date === null ? null : [wName("date"), comment.date],
    comment.initials === null ? null : [wName("initials"), comment.initials],
  ];
  const paraId = keyedEntry(comment, arrivedKeyed) ? comment.paraId : null;
  const body = renderCommentBody(comment.text, paraId);
  return elementXml(
    wName("comment"),
    attrs.filter((attr): attr is XmlAttr => attr !== null),
    [body]
  );
}

/**
 * What a part carrying comments this editor wrote has to declare. Every entry it writes is spelled
 * under `w`, so the root binds it rather than each entry declaring it again.
 */
const COMMENT_MARKUP: RootDeclarations = { namespaces: { w: NAMESPACES.w } };

/**
 * The same, for a part carrying a thread key: the key is a `w14:paraId`, and a reader that does
 * not know that namespace is told it may pass over it.
 */
const THREAD_MARKUP: RootDeclarations = {
  namespaces: { w: NAMESPACES.w, w14: NAMESPACES.w14, mc: NAMESPACES.mc },
  ignorable: ["w14"],
};

export function currentCommentBodies(
  references: ReadonlyMap<string, CommentReferenceData>
): ReadonlyMap<string, CommentReferenceData | CommentReplyData> {
  const comments = new Map<string, CommentReferenceData | CommentReplyData>();
  for (const [id, comment] of references) {
    comments.set(id, comment);
    for (const reply of comment.replies) comments.set(reply.id, reply);
  }
  return comments;
}

function originalThreadIds(
  comments: ImportedComments,
  rootIds: ReadonlySet<string>
): ReadonlySet<string> {
  const ids = new Set(rootIds);
  const pending = Array.from(rootIds);
  while (pending.length > 0) {
    const parent = pending.shift();
    if (parent === undefined) break;
    for (const reply of comments.repliesByParentId.get(parent) ?? []) {
      if (ids.has(reply.id)) continue;
      ids.add(reply.id);
      pending.push(reply.id);
    }
  }
  return ids;
}

function commentsXml(
  references: ReadonlyMap<string, CommentReferenceData>,
  comments: ImportedComments,
  originallyReferenced: ReadonlySet<string>
): string {
  const currentBodies = currentCommentBodies(references);
  const arrivedKeyed = new Set(
    Array.from(comments.byId.values()).flatMap((entry) =>
      entry.paraId === null ? [] : [entry.id]
    )
  );
  const hasThreadMetadata = Array.from(currentBodies.values()).some((comment) =>
    keyedEntry(comment, arrivedKeyed)
  );
  const arrived = arrivedEntries(comments.xml);
  const originalThreads = originalThreadIds(comments, originallyReferenced);
  const pieces: string[] = [];
  const written = new Set<string>();

  for (const original of comments.ordered) {
    const current = currentBodies.get(original.id);
    if (current) {
      pieces.push(renderedComment(current, arrivedKeyed, arrived));
      written.add(original.id);
      continue;
    }
    // An orphan was not deleted through the editor and stays untouched.
    if (!originalThreads.has(original.id)) {
      pieces.push(original.xml);
      written.add(original.id);
    }
  }
  for (const [id, comment] of currentBodies) {
    if (!written.has(id)) {
      pieces.push(renderedComment(comment, arrivedKeyed, arrived));
    }
  }

  if (comments.xml === null) {
    const compatibility = hasThreadMetadata
      ? ` ${xmlnsDecl("w14")} ${xmlnsDecl("mc")} mc:Ignorable="w14"`
      : "";
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:comments ${xmlnsDecl("w")}${compatibility}>${pieces.join("")}</w:comments>`
    );
  }

  const rewritten = splicePart(comments.xml, {
    root: COMMENTS_ROOT,
    replaceChildren: pieces.join(""),
  });
  return ensureRootDeclarations(
    rewritten,
    hasThreadMetadata ? THREAD_MARKUP : COMMENT_MARKUP
  );
}

/**
 * Plans the Comments part, relationship and content type only when comment state changed, and the
 * people part beside them for an author whose identity the document has yet to record.
 */
function planCommentParts(
  doc: PMNode,
  session: SessionStore,
  context: PartPlanContext
): ReadonlyMap<string, Uint8Array> | null {
  const bodyChanged = commentsChanged(doc, session);
  const threadChanged = extensionsChanged(doc, session);
  if (!bodyChanged && !threadChanged) return null;

  const { relationships, contentTypes } = context;
  const references = commentReferencesIn(doc);
  const addingPart = commentsPart.pathIn(session) === null;
  const partPath = commentsPart.writePathIn(session);
  const parts = new Map<string, Uint8Array>();

  if (addingPart) {
    const target = partPath.slice(directoryOf(session.mainPartPath).length);
    relationships.add({ type: commentsPart.relType, target });
  }

  if (bodyChanged) {
    parts.set(
      partPath,
      encodeUtf8(
        commentsXml(references, session.comments, session.commentReferenceIds),
        session.comments.hadBom
      )
    );
  }

  if (addingPart || session.comments.xml === null) {
    contentTypes.addOverride(partPath, commentsPart.contentType);
  }

  if (
    threadChanged &&
    (references.size > 0 || session.comments.extendedPartPath !== null)
  ) {
    const addingExtendedPart = commentsExtendedPart.pathIn(session) === null;
    const extendedPartPath = commentsExtendedPart.writePathIn(session);
    if (addingExtendedPart) {
      const target = extendedPartPath.slice(
        directoryOf(session.mainPartPath).length
      );
      relationships.add({ type: commentsExtendedPart.relType, target });
    }
    parts.set(
      extendedPartPath,
      encodeUtf8(
        extensionsXml(
          references,
          session.comments,
          session.commentReferenceIds
        ),
        session.comments.extendedHadBom
      )
    );
    if (addingExtendedPart || session.comments.extendedXml === null) {
      contentTypes.addOverride(
        extendedPartPath,
        commentsExtendedPart.contentType
      );
    }
  }
  if (bodyChanged) {
    const people = planPeoplePart(
      peoplePart,
      currentCommentBodies(references).values(),
      session,
      context
    );
    for (const [path, bytes] of people ?? []) parts.set(path, bytes);
  }
  return parts;
}

export const commentsPlanner: PartPlanner = {
  name: "comments",
  plan: planCommentParts,
};
