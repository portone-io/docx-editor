/** Reads comment threads and anchors from the editor document. */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import type { CommentReplyData } from "../../../docx/comments";
import { commentOwned } from "../../../schema/protection";
import { protectionOf } from "../../../schema/protectionState";
import type { DocumentComment } from "./model";

interface MarkerPositions {
  starts: Map<string, number>;
  ends: Map<string, number>;
  references: Array<{ node: PMNode; pos: number }>;
}

export function stringAttr(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

export function repliesAttr(value: unknown): readonly CommentReplyData[] {
  return Array.isArray(value) ? (value as CommentReplyData[]) : [];
}

function markerPositions(doc: PMNode): MarkerPositions {
  const starts = new Map<string, number>();
  const ends = new Map<string, number>();
  const references: Array<{ node: PMNode; pos: number }> = [];
  doc.descendants((node, pos) => {
    const id = stringAttr(node.attrs.id);
    if (id === null) return true;
    if (node.type.name === "commentStart" && !starts.has(id)) {
      starts.set(id, pos);
    } else if (node.type.name === "commentEnd" && !ends.has(id)) {
      ends.set(id, pos);
    } else if (node.type.name === "commentReference") {
      references.push({ node, pos });
    }
    return true;
  });
  return { starts, ends, references };
}

/** Comments in document order, including point comments that have no explicit range. */
export function documentComments(
  state: EditorState
): readonly DocumentComment[] {
  const markers = markerPositions(state.doc);
  return markers.references.map(({ node, pos }) => {
    const id = stringAttr(node.attrs.id) ?? "";
    const start = markers.starts.get(id);
    const end = markers.ends.get(id);
    const hasRange = start !== undefined && end !== undefined && start < end;
    const point = start ?? end ?? pos;
    return {
      id,
      author: stringAttr(node.attrs.author),
      authorId: stringAttr(node.attrs.authorId),
      initials: stringAttr(node.attrs.initials),
      date: stringAttr(node.attrs.date),
      text: stringAttr(node.attrs.text) ?? "",
      from: hasRange ? start + 1 : point,
      to: hasRange ? end : point,
      referencePos: pos,
      resolved: node.attrs.resolved === true,
      replies: repliesAttr(node.attrs.replies).map((reply) => ({
        id: reply.id,
        author: reply.author,
        authorId: reply.authorId,
        initials: reply.initials,
        date: reply.date,
        text: reply.text,
      })),
    };
  });
}

/**
 * Whether the body of this comment, or of one of its replies, may be edited or deleted here: the
 * protection lets comments be edited at all, and the body is one of one's own to edit
 * (`schema/protection`). Replying and settling a thread are not governed by this; a comment command
 * asked without `dispatch` answers for those.
 */
export function canEditComment(
  state: EditorState,
  commentId: string,
  replyId: string | null = null
): boolean {
  const rules = protectionOf(state);
  if (rules.protection === "readOnly") return false;
  const comment = documentComments(state).find(
    (entry) => entry.id === commentId
  );
  const body =
    replyId === null
      ? comment
      : comment?.replies.find((reply) => reply.id === replyId);
  return body !== undefined && commentOwned(rules, body.authorId);
}
