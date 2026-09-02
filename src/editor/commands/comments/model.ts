/** Public comment values and shared attribute readers. */

export interface CommentAuthor {
  /**
   * The identity behind the name, an opaque string the host application chooses. Recorded in the
   * document's people part, and what decides whose comment a comment is (`schema/protection`).
   */
  id?: string;
  name: string;
  initials?: string;
}

export interface NewComment {
  text: string;
  author: string;
  initials?: string;
  /** ISO 8601 timestamp. The current time is used when omitted. */
  date?: string;
}

export interface DocumentComment {
  id: string;
  author: string | null;
  initials: string | null;
  date: string | null;
  text: string;
  from: number;
  to: number;
  referencePos: number;
  resolved: boolean;
  replies: readonly DocumentCommentReply[];
}

export interface DocumentCommentReply {
  id: string;
  author: string | null;
  initials: string | null;
  date: string | null;
  text: string;
}
