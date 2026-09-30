/**
 * Maps imported comment threads to and from ProseMirror reference data.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { ImportedComments } from "./reading";

/**
 * A comment as the reference node in the story carries it: who wrote it, where its thread state
 * hangs, and the replies under it. What it says is not here - a body is a story of its own, held
 * on the document node under `comment:<id>` (`docx/story`).
 */
export interface CommentReferenceData {
  id: string;
  author: string | null;
  authorId: string | null;
  initials: string | null;
  /** As the file writes it: the author's wall clock (`./dates`) */
  date: string | null;
  paraId: string;
  resolved: boolean;
  extensionXml: string | null;
  threadImported: boolean;
  replies: readonly CommentReplyData[];
  durableId: string | null;
  dateUtc: string | null;
}

export interface CommentReplyData {
  id: string;
  author: string | null;
  authorId: string | null;
  initials: string | null;
  date: string | null;
  paraId: string;
  parentParaId: string;
  extensionXml: string | null;
  durableId: string | null;
  dateUtc: string | null;
}

/** A stable eight-digit hex number below 0x80000000, hashed from the seed */
export function seededHexId(seed: string): string {
  let hash = 0x811c9dc5;
  for (const char of seed) {
    hash ^= char.charCodeAt(0);
    hash = Math.imul(hash, 0x01000193);
  }
  const value = (hash >>> 0) & 0x7fffffff;
  return (value === 0 ? 1 : value).toString(16).toUpperCase().padStart(8, "0");
}

/** Replies below one referenced comment, in the order of the Comments part. */
export function importedCommentReplies(
  comments: ImportedComments,
  rootId: string
): readonly CommentReplyData[] {
  const replies: CommentReplyData[] = [];
  const visited = new Set([rootId]);
  const pending = [...(comments.repliesByParentId.get(rootId) ?? [])]
    .reverse()
    .map((reply) => ({ parentId: rootId, reply }));
  while (pending.length > 0) {
    const item = pending.pop();
    if (item === undefined || visited.has(item.reply.id)) continue;
    const { parentId, reply } = item;
    visited.add(reply.id);
    const parent = comments.byId.get(parentId);
    const paraId = reply.paraId ?? seededHexId(`comment-${reply.id}`);
    const parentParaId =
      reply.parentParaId ??
      parent?.paraId ??
      seededHexId(`comment-${parentId}`);
    replies.push({
      id: reply.id,
      author: reply.author,
      authorId: reply.authorId,
      initials: reply.initials,
      date: reply.date,
      paraId,
      parentParaId,
      extensionXml: reply.extensionXml,
      durableId: reply.durableId,
      dateUtc: reply.dateUtc,
    });
    const children = comments.repliesByParentId.get(reply.id) ?? [];
    for (let index = children.length - 1; index >= 0; index -= 1) {
      pending.push({ parentId: reply.id, reply: children[index] });
    }
  }
  return replies;
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

/** These comments and every reply under them, however deep, in the part as it arrived */
export function originalThreadIds(
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

function nullableString(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function replyData(value: unknown): CommentReplyData[] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((candidate) => {
    if (typeof candidate !== "object" || candidate === null) return [];
    const entry = candidate as Record<string, unknown>;
    const id = nullableString(entry.id);
    const paraId = nullableString(entry.paraId);
    const parentParaId = nullableString(entry.parentParaId);
    if (id === null || paraId === null || parentParaId === null) return [];
    return [
      {
        id,
        author: nullableString(entry.author),
        authorId: nullableString(entry.authorId),
        initials: nullableString(entry.initials),
        date: nullableString(entry.date),
        paraId,
        parentParaId,
        extensionXml: nullableString(entry.extensionXml),
        durableId: nullableString(entry.durableId),
        dateUtc: nullableString(entry.dateUtc),
      },
    ];
  });
}

function referenceData(node: PMNode): CommentReferenceData | null {
  if (node.type.name !== "commentReference") return null;
  const id = nullableString(node.attrs.id);
  if (id === null) return null;
  return {
    id,
    author: nullableString(node.attrs.author),
    authorId: nullableString(node.attrs.authorId),
    initials: nullableString(node.attrs.initials),
    date: nullableString(node.attrs.date),
    paraId: nullableString(node.attrs.paraId) ?? "00000001",
    resolved: node.attrs.resolved === true,
    extensionXml: nullableString(node.attrs.extensionXml),
    threadImported: node.attrs.threadImported === true,
    replies: replyData(node.attrs.replies),
    durableId: nullableString(node.attrs.durableId),
    dateUtc: nullableString(node.attrs.dateUtc),
  };
}

/** The first reference for each comment id, in document order. */
export function commentReferencesIn(
  doc: PMNode
): ReadonlyMap<string, CommentReferenceData> {
  const references = new Map<string, CommentReferenceData>();
  doc.descendants((node) => {
    const comment = referenceData(node);
    if (comment && !references.has(comment.id)) {
      references.set(comment.id, comment);
    }
    return true;
  });
  return references;
}
