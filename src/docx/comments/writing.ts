/**
 * Plans changed comment package parts while preserving untouched XML.
 */

import type { Node as PMNode } from "prosemirror-model";
import { elementXml, type XmlAttr, xmlnsAttr } from "../../ooxml/element";
import { DocxExportError } from "../../ooxml/errors";
import { wName, xmlnsDecl } from "../../ooxml/names";
import { encodeUtf8 } from "../../ooxml/xml";
import { CONTENT_TYPES_PATH } from "../packageParts";
import { directoryOf, type RelationshipWriter } from "../relationships";
import type { SessionStore } from "../session";
import { withContentType } from "./contentTypes";
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
  const root = partRoot(comments.extendedXml, EXTENSIONS_ROOT);
  return withinRoot(
    comments.extendedXml,
    root,
    openedTag(root),
    pieces.join("")
  );
}

/** The root element a comment part is written around, and what the part is called when it has none */
interface PartRootName {
  localName: string;
  partName: string;
}

const COMMENTS_ROOT: PartRootName = {
  localName: "comments",
  partName: "Comments",
};

const EXTENSIONS_ROOT: PartRootName = {
  localName: "commentsEx",
  partName: "Comments Extended",
};

/** Where the root element of a part stands: its opening tag as written, and where its closing tag begins */
interface PartRoot {
  openAt: number;
  openTag: string;
  name: string;
  /** null for an empty element, which has no closing tag until it is opened */
  closeAt: number | null;
}

/**
 * Reads the root of a comment part, or says why the part cannot be rewritten around it.
 *
 * The same question stands in the export invariant list, so a refusal here is one a caller could
 * have read there first.
 */
function readPartRoot(
  xml: string,
  { localName, partName }: PartRootName
): { root: PartRoot } | { problem: string } {
  const open = new RegExp(`<(?:[^\\s<>/:="']+:)?${localName}\\b[^>]*>`).exec(
    xml
  );
  const name = open === null ? undefined : /^<([^\s>/]+)/.exec(open[0])?.[1];
  if (open === null || name === undefined) {
    return { problem: `the ${partName} part has no ${localName} root element` };
  }
  if (open[0].endsWith("/>")) {
    return {
      root: { openAt: open.index, openTag: open[0], name, closeAt: null },
    };
  }
  const closeAt = xml.lastIndexOf(`</${name}>`);
  if (closeAt === -1) {
    return { problem: `the ${partName} part has no closing ${localName} tag` };
  }
  return { root: { openAt: open.index, openTag: open[0], name, closeAt } };
}

function partRoot(xml: string, rootName: PartRootName): PartRoot {
  const reading = readPartRoot(xml, rootName);
  if ("problem" in reading) {
    throw new DocxExportError("malformed-xml", reading.problem);
  }
  return reading.root;
}

/** Why the Comments part cannot be rewritten around its root, or null when it can */
export function commentsRootProblem(xml: string): string | null {
  const reading = readPartRoot(xml, COMMENTS_ROOT);
  return "problem" in reading ? reading.problem : null;
}

/** Why the extended part cannot be rewritten around its root, or null when it can */
export function extensionsRootProblem(xml: string): string | null {
  const reading = readPartRoot(xml, EXTENSIONS_ROOT);
  return "problem" in reading ? reading.problem : null;
}

/** The opening tag with something to stand inside it: an empty element is opened before an entry can go in */
function openedTag(root: PartRoot): string {
  return root.closeAt === null ? `${root.openTag.slice(0, -2)}>` : root.openTag;
}

/** The part with the entries written inside its root and everything around the root standing as it came */
function withinRoot(
  xml: string,
  root: PartRoot,
  openTag: string,
  entries: string
): string {
  const head = xml.slice(0, root.openAt) + openTag + entries;
  return root.closeAt === null
    ? `${head}</${root.name}>${xml.slice(root.openAt + root.openTag.length)}`
    : head + xml.slice(root.closeAt);
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
    xmlnsAttr("w"),
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

function withThreadMarkupCompatibility(openTag: string): string {
  let updated = openTag;
  if (!/\sxmlns:w14\s*=/.test(updated)) {
    updated = updated.replace(/>$/, ` ${xmlnsDecl("w14")}>`);
  }
  if (!/\sxmlns:mc\s*=/.test(updated)) {
    updated = updated.replace(/>$/, ` ${xmlnsDecl("mc")}>`);
  }
  const ignorable = /\smc:Ignorable\s*=\s*(["'])([^"']*)\1/.exec(updated);
  if (!ignorable) {
    return updated.replace(/>$/, ' mc:Ignorable="w14">');
  }
  const tokens = ignorable[2].split(/\s+/).filter(Boolean);
  if (tokens.includes("w14")) return updated;
  const replacement = ` mc:Ignorable=${ignorable[1]}${[...tokens, "w14"].join(" ")}${ignorable[1]}`;
  return (
    updated.slice(0, ignorable.index) +
    replacement +
    updated.slice(ignorable.index + ignorable[0].length)
  );
}

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

  const root = partRoot(comments.xml, COMMENTS_ROOT);
  const openTag = hasThreadMetadata
    ? withThreadMarkupCompatibility(openedTag(root))
    : openedTag(root);
  return withinRoot(comments.xml, root, openTag, pieces.join(""));
}

export interface CommentPartChanges {
  parts: ReadonlyMap<string, Uint8Array>;
}

/**
 * Plans the Comments part, relationship and content type only when comment state changed, and the
 * people part beside them for an author whose identity the document has yet to record.
 */
export function planCommentParts(
  doc: PMNode,
  session: SessionStore,
  relationships: RelationshipWriter,
  currentContentTypes?: Uint8Array
): CommentPartChanges | null {
  const bodyChanged = commentsChanged(doc, session);
  const threadChanged = extensionsChanged(doc, session);
  if (!bodyChanged && !threadChanged) return null;

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
    const contentTypes = withContentType(
      session.parts,
      partPath,
      commentsPart.contentType,
      currentContentTypes
    );
    if (contentTypes) parts.set(CONTENT_TYPES_PATH, contentTypes);
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
      const contentTypes = withContentType(
        session.parts,
        extendedPartPath,
        commentsExtendedPart.contentType,
        parts.get(CONTENT_TYPES_PATH) ?? currentContentTypes
      );
      if (contentTypes) parts.set(CONTENT_TYPES_PATH, contentTypes);
    }
  }
  if (bodyChanged) {
    const people = planPeoplePart(
      peoplePart,
      currentCommentBodies(references).values(),
      session,
      relationships,
      parts.get(CONTENT_TYPES_PATH) ?? currentContentTypes
    );
    for (const [path, bytes] of people ?? []) parts.set(path, bytes);
  }
  return { parts };
}
