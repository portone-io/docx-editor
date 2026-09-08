/** Reads comment threads and anchors from the editor document. */

import type { EditorState } from "prosemirror-state";
import { commentOwned } from "../../../schema/protection";
import { protectionOf } from "../../../schema/protectionState";
import { commentProjection } from "../../plugins/commentDecorations";
import type { DocumentComment } from "./model";

/** Comments in document order, including point comments that have no explicit range. */
export function documentComments(
  state: EditorState
): readonly DocumentComment[] {
  // Public records have mutable fields. Keep caller edits out of the shared projection and the
  // command lookup, including edits to a reply's fields.
  return commentProjection.read(state).comments.map((comment) => ({
    ...comment,
    replies: comment.replies.map((reply) => ({ ...reply })),
  }));
}

/**
 * The comment this id names, or undefined where the document holds none. Answered off the lookup
 * the projection built, so a command asking after one comment does not read the whole list.
 */
export function commentById(
  state: EditorState,
  id: string
): DocumentComment | undefined {
  return commentProjection.read(state).byId.get(id);
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
  const comment = commentById(state, commentId);
  const body =
    replyId === null
      ? comment
      : comment?.replies.find((reply) => reply.id === replyId);
  return body !== undefined && commentOwned(rules, body.authorId);
}
