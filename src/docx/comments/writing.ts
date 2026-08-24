/**
 * Plans changed comment package parts while preserving untouched XML.
 */

import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../../ooxml/errors";
import { decodeUtf8, encodeUtf8, escapeXml, W_NS } from "../../ooxml/xml";
import { directoryOf, type RelationshipWriter } from "../relationships";
import type { SessionStore } from "../session";
import {
  COMMENTS_CONTENT_TYPE,
  COMMENTS_EXTENDED_CONTENT_TYPE,
  COMMENTS_EXTENDED_REL_TYPE,
  COMMENTS_REL_TYPE,
  CONTENT_TYPES_PATH,
  MC_NS,
  W14_NS,
  W15_NS,
} from "./constants";
import {
  type CommentReferenceData,
  type CommentReplyData,
  commentReferencesIn,
} from "./model";
import type { ImportedComments } from "./reading";

function commentsChanged(doc: PMNode, session: SessionStore): boolean {
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
  return false;
}

function renderedExtension(
  comment: CommentReferenceData | CommentReplyData
): string {
  if (comment.extensionXml !== null) return comment.extensionXml;
  const parent =
    "parentParaId" in comment
      ? ` w15:paraIdParent="${escapeXml(comment.parentParaId)}"`
      : "";
  const done =
    "resolved" in comment ? ` w15:done="${comment.resolved ? "1" : "0"}"` : "";
  return `<w15:commentEx w15:paraId="${escapeXml(comment.paraId)}"${parent}${done}/>`;
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
      pieces.push(renderedExtension(item));
      written.add(id);
    } else if (id === undefined || !originalThreads.has(id)) {
      pieces.push(original.xml);
    }
  }
  for (const [id, comment] of current) {
    if (!written.has(id)) pieces.push(renderedExtension(comment));
  }

  if (comments.extendedXml === null) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w15:commentsEx xmlns:w15="${W15_NS}">${pieces.join("")}</w15:commentsEx>`
    );
  }
  const open = /<(?:[\w.-]+:)?commentsEx\b[^>]*>/.exec(comments.extendedXml);
  if (!open) {
    throw new DocxExportError(
      "malformed-xml",
      "the Comments Extended part has no commentsEx root element"
    );
  }
  const name = /^<([^\s>]+)/.exec(open[0])?.[1];
  const close = name ? comments.extendedXml.lastIndexOf(`</${name}>`) : -1;
  if (close === -1) {
    throw new DocxExportError(
      "malformed-xml",
      "the Comments Extended part has no closing commentsEx tag"
    );
  }
  const bodyAt = open.index + open[0].length;
  return (
    comments.extendedXml.slice(0, bodyAt) +
    pieces.join("") +
    comments.extendedXml.slice(close)
  );
}

function extensionsChanged(doc: PMNode, session: SessionStore): boolean {
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

function commentTextXml(text: string): string {
  const lines = text.split("\n");
  const pieces: string[] = [];
  lines.forEach((line, index) => {
    if (index > 0) pieces.push("<w:br/>");
    if (line.length > 0 || lines.length === 1) {
      pieces.push(`<w:t xml:space="preserve">${escapeXml(line)}</w:t>`);
    }
  });
  return `<w:r>${pieces.join("")}</w:r>`;
}

function carriesThreadMetadata(
  comment: CommentReferenceData | CommentReplyData
): boolean {
  return (
    comment.extensionXml !== null ||
    ("threadImported" in comment && !comment.threadImported) ||
    "parentParaId" in comment
  );
}

function renderedComment(
  comment: CommentReferenceData | CommentReplyData
): string {
  if (comment.imported && comment.commentXml !== null) {
    return comment.commentXml;
  }
  const attrs = [
    `w:id="${escapeXml(comment.id)}"`,
    comment.author === null ? null : `w:author="${escapeXml(comment.author)}"`,
    comment.date === null ? null : `w:date="${escapeXml(comment.date)}"`,
    comment.initials === null
      ? null
      : `w:initials="${escapeXml(comment.initials)}"`,
  ]
    .filter((entry): entry is string => entry !== null)
    .join(" ");
  const threaded = carriesThreadMetadata(comment);
  const paragraphAttrs = threaded
    ? ` xmlns:w14="${W14_NS}" w14:paraId="${escapeXml(comment.paraId)}"`
    : "";
  return `<w:comment xmlns:w="${W_NS}" ${attrs}><w:p${paragraphAttrs}>${commentTextXml(comment.text)}</w:p></w:comment>`;
}

function withThreadMarkupCompatibility(openTag: string): string {
  let updated = openTag;
  if (!/\sxmlns:w14\s*=/.test(updated)) {
    updated = updated.replace(/>$/, ` xmlns:w14="${W14_NS}">`);
  }
  if (!/\sxmlns:mc\s*=/.test(updated)) {
    updated = updated.replace(/>$/, ` xmlns:mc="${MC_NS}">`);
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

function currentCommentBodies(
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
  const hasThreadMetadata = Array.from(currentBodies.values()).some(
    carriesThreadMetadata
  );
  const originalThreads = originalThreadIds(comments, originallyReferenced);
  const pieces: string[] = [];
  const written = new Set<string>();

  for (const original of comments.ordered) {
    const current = currentBodies.get(original.id);
    if (current) {
      pieces.push(renderedComment(current));
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
    if (!written.has(id)) pieces.push(renderedComment(comment));
  }

  if (comments.xml === null) {
    const compatibility = hasThreadMetadata
      ? ` xmlns:w14="${W14_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="w14"`
      : "";
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:comments xmlns:w="${W_NS}"${compatibility}>${pieces.join("")}</w:comments>`
    );
  }

  const open = /<(?:[\w.-]+:)?comments\b[^>]*>/.exec(comments.xml);
  if (!open) {
    throw new DocxExportError(
      "malformed-xml",
      "the Comments part has no comments root element"
    );
  }
  const name = /^<([^\s>]+)/.exec(open[0])?.[1];
  const close = name ? comments.xml.lastIndexOf(`</${name}>`) : -1;
  if (close === -1) {
    throw new DocxExportError(
      "malformed-xml",
      "the Comments part has no closing comments tag"
    );
  }
  const openTag = hasThreadMetadata
    ? withThreadMarkupCompatibility(open[0])
    : open[0];
  return (
    comments.xml.slice(0, open.index) +
    openTag +
    pieces.join("") +
    comments.xml.slice(close)
  );
}

const TYPES_OPEN_TAG = /<(?:[\w.-]+:)?Types\b[^>]*>/;

function withContentType(
  parts: Map<string, Uint8Array>,
  partPath: string,
  contentType: string,
  current: Uint8Array | undefined
): Uint8Array | null {
  const original = current ?? parts.get(CONTENT_TYPES_PATH);
  if (!original) {
    throw new DocxExportError(
      "missing-content-types",
      `cannot add a Comments part to a package that has no ${CONTENT_TYPES_PATH}`
    );
  }
  const { text, hadBom } = decodeUtf8(original);
  const partName = `/${partPath}`;
  if (
    new RegExp(
      `<(?:[\\w.-]+:)?Override[^>]+PartName=["']${partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}["']`,
      "i"
    ).test(text)
  ) {
    return null;
  }
  const open = TYPES_OPEN_TAG.exec(text);
  if (!open) {
    throw new DocxExportError(
      "malformed-xml",
      `${CONTENT_TYPES_PATH} has no Types element`
    );
  }
  const rootName = /^<([^\s>]+)/.exec(open[0])?.[1] ?? "Types";
  const separator = rootName.indexOf(":");
  const prefix = separator < 0 ? "" : `${rootName.slice(0, separator)}:`;
  const declaration = `<${prefix}Override PartName="${partName}" ContentType="${contentType}"/>`;
  const at = open.index + open[0].length;
  return encodeUtf8(text.slice(0, at) + declaration + text.slice(at), hadBom);
}

function availableCommentsPath(session: SessionStore): string {
  const directory = directoryOf(session.mainPartPath);
  for (let suffix = 0; ; suffix += 1) {
    const name = suffix === 0 ? "comments.xml" : `comments${suffix + 1}.xml`;
    const path = directory + name;
    if (!session.parts.has(path)) return path;
  }
}

function availableCommentsExtendedPath(session: SessionStore): string {
  const directory = directoryOf(session.mainPartPath);
  for (let suffix = 0; ; suffix += 1) {
    const name =
      suffix === 0
        ? "commentsExtended.xml"
        : `commentsExtended${suffix + 1}.xml`;
    const path = directory + name;
    if (!session.parts.has(path)) return path;
  }
}

export interface CommentPartChanges {
  parts: ReadonlyMap<string, Uint8Array>;
}

/** Plans the Comments part, relationship and content type only when comment state changed. */
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
  const addingPart = session.comments.partPath === null;
  const partPath = session.comments.partPath ?? availableCommentsPath(session);
  const parts = new Map<string, Uint8Array>();

  if (addingPart) {
    const target = partPath.slice(directoryOf(session.mainPartPath).length);
    relationships.add({ type: COMMENTS_REL_TYPE, target });
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
      COMMENTS_CONTENT_TYPE,
      currentContentTypes
    );
    if (contentTypes) parts.set(CONTENT_TYPES_PATH, contentTypes);
  }

  if (
    threadChanged &&
    (references.size > 0 || session.comments.extendedPartPath !== null)
  ) {
    const addingExtendedPart = session.comments.extendedPartPath === null;
    const extendedPartPath =
      session.comments.extendedPartPath ??
      availableCommentsExtendedPath(session);
    if (addingExtendedPart) {
      const target = extendedPartPath.slice(
        directoryOf(session.mainPartPath).length
      );
      relationships.add({ type: COMMENTS_EXTENDED_REL_TYPE, target });
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
        COMMENTS_EXTENDED_CONTENT_TYPE,
        parts.get(CONTENT_TYPES_PATH) ?? currentContentTypes
      );
      if (contentTypes) parts.set(CONTENT_TYPES_PATH, contentTypes);
    }
  }
  return { parts };
}
