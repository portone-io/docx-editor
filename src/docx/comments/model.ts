/**
 * Maps imported comment threads to and from ProseMirror reference data.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { ImportedComments } from "./reading";

export interface CommentReferenceData {
  id: string;
  author: string | null;
  initials: string | null;
  date: string | null;
  text: string;
  commentXml: string | null;
  imported: boolean;
  paraId: string;
  resolved: boolean;
  extensionXml: string | null;
  threadImported: boolean;
  replies: readonly CommentReplyData[];
}

export interface CommentReplyData {
  id: string;
  author: string | null;
  initials: string | null;
  date: string | null;
  text: string;
  commentXml: string | null;
  imported: boolean;
  paraId: string;
  parentParaId: string;
  extensionXml: string | null;
}

/** A stable Word paragraph id for a comment created or upgraded by the editor. */
export function commentParaId(seed: string): string {
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
    const paraId = reply.paraId ?? commentParaId(`comment-${reply.id}`);
    const parentParaId =
      reply.parentParaId ??
      parent?.paraId ??
      commentParaId(`comment-${parentId}`);
    replies.push({
      id: reply.id,
      author: reply.author,
      initials: reply.initials,
      date: reply.date,
      text: reply.text,
      commentXml: reply.xml,
      imported: true,
      paraId,
      parentParaId,
      extensionXml: reply.extensionXml,
    });
    const children = comments.repliesByParentId.get(reply.id) ?? [];
    for (let index = children.length - 1; index >= 0; index -= 1) {
      pending.push({ parentId: reply.id, reply: children[index] });
    }
  }
  return replies;
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
        initials: nullableString(entry.initials),
        date: nullableString(entry.date),
        text: nullableString(entry.text) ?? "",
        commentXml: nullableString(entry.commentXml),
        imported: entry.imported === true,
        paraId,
        parentParaId,
        extensionXml: nullableString(entry.extensionXml),
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
    initials: nullableString(node.attrs.initials),
    date: nullableString(node.attrs.date),
    text: nullableString(node.attrs.text) ?? "",
    commentXml: nullableString(node.attrs.commentXml),
    imported: node.attrs.imported === true,
    paraId: nullableString(node.attrs.paraId) ?? "00000001",
    resolved: node.attrs.resolved === true,
    extensionXml: nullableString(node.attrs.extensionXml),
    threadImported: node.attrs.threadImported === true,
    replies: replyData(node.attrs.replies),
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
