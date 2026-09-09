/** Public comment values and shared attribute readers. */

import type { CommentReplyData } from "../../../docx/comments";

/** Who a comment is written by: the name it is shown under, and the identity behind it. */
export interface CommentAuthor {
  /**
   * The identity behind the name, an opaque string the host application chooses. Recorded in the
   * document's people part, and what decides whose comment a comment is (`schema/protection`).
   */
  id: string;
  name: string;
  initials?: string;
}

/**
 * A stretch of text a comment is written for: a pair of document positions, the way a comment
 * already in the document reports its own (`DocumentComment.from`, `.to`).
 */
export interface CommentRange {
  from: number;
  to: number;
}

export interface NewComment {
  text: string;
  author: string;
  /** The identity behind `author`. A comment written without one belongs to nobody in particular */
  authorId?: string;
  initials?: string;
  /** ISO 8601 timestamp. The current time is used when omitted. */
  date?: string;
}

/**
 * One comment as a reader of the document sees it. Every field is readonly: the record is the
 * editor's own, worked out once per edit and handed to whoever asks (`editor/plugins/documentProjection`).
 */
export interface DocumentComment {
  readonly id: string;
  readonly author: string | null;
  readonly authorId: string | null;
  readonly initials: string | null;
  readonly date: string | null;
  /** What it says, as plain text. The formatting behind it is `setCommentBody`'s to write */
  readonly text: string;
  readonly from: number;
  readonly to: number;
  readonly referencePos: number;
  /**
   * Whether the comment still marks a stretch of the document: both range markers stand, and the
   * opening one stands before the closing one. A comment that lost both markers with the text an
   * edit deleted is detached - `from` and `to` are then both where its reference stands - and is
   * drawn nowhere on the page.
   */
  readonly anchored: boolean;
  readonly resolved: boolean;
  readonly replies: readonly DocumentCommentReply[];
}

export interface DocumentCommentReply {
  readonly id: string;
  readonly author: string | null;
  readonly authorId: string | null;
  readonly initials: string | null;
  readonly date: string | null;
  readonly text: string;
}

export function stringAttr(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

export function repliesAttr(value: unknown): readonly CommentReplyData[] {
  return Array.isArray(value) ? (value as CommentReplyData[]) : [];
}
