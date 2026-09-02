/**
 * The protection an editor runs under: what the document as a whole may receive, over and above
 * the locks its own controls carry (`./locks`).
 *
 * The levels are named after OOXML `ST_DocProtect`, the enumeration `w:documentProtection/@w:edit`
 * takes (ECMA-376 Part 1 §17.15.1.29): `readOnly` takes no edit at all, `comments` takes comments
 * and nothing else, `none` takes every body edit. No level lifts the ownership rule over comments:
 * whose comment may be edited is `editableComments`' question at every level, `none` included.
 * `trackedChanges` and `forms`, the two values of the enumeration not modelled yet, would be added
 * here.
 *
 * The protection is editor state rather than document state: it is the reader's standing, not the
 * file's, so nothing here is written to or read from OOXML. `editor/plugins/documentProtection`
 * holds it in the state and `./protectionState` reads it back; the guard (`transactionAllowed` in
 * `./locks`) asks the judgements here. They take documents alone, so that the `./core` entry can
 * make the same judgement over a file without an editor.
 */

import { Fragment, type Node as PMNode } from "prosemirror-model";

export type EditingProtection = "none" | "readOnly" | "comments";

/**
 * Whose comments may be edited or deleted. `own` is a comment carrying no recognised identity, or
 * the very identity comments are being written under; `all` is every comment there is, which is
 * the standing of a moderator.
 */
export type EditableComments = "own" | "all";

export interface ProtectionState {
  protection: EditingProtection;
  /** The identity comments are written under. Null where nobody is writing any, a reader above all */
  authorId: string | null;
  editableComments: EditableComments;
}

const COMMENT_NODES = new Set([
  "commentStart",
  "commentEnd",
  "commentReference",
]);

/** Whether this node is one of the three a comment stands in the story as */
export function isCommentNode(node: PMNode): boolean {
  return COMMENT_NODES.has(node.type.name);
}

/**
 * The document with every comment taken out of it.
 *
 * `Fragment.fromArray` joins the text a marker had split, so a document a comment was put into
 * reads the same as it did before the comment, and the two compare equal.
 */
function withoutComments(node: PMNode): PMNode {
  if (node.isLeaf) return node;
  const kept: PMNode[] = [];
  node.forEach((child) => {
    if (!isCommentNode(child)) kept.push(withoutComments(child));
  });
  return node.copy(Fragment.fromArray(kept));
}

/** Whether the two documents differ in nothing but their comments */
export function changesOnlyComments(before: PMNode, after: PMNode): boolean {
  return withoutComments(before).eq(withoutComments(after));
}

/**
 * A comment as ownership is judged on it: who wrote it, and everything the author of that comment
 * may rewrite.
 *
 * The identity is part of the body so that a rewrite of it can be told from a rewrite of what it
 * says. It is the one part nobody rewrites, its own author included.
 */
interface CommentBody {
  authorId: string | null;
  author: string | null;
  initials: string | null;
  date: string | null;
  text: string;
}

interface CommentThread extends CommentBody {
  replies: ReadonlyMap<string, CommentBody>;
}

function stringOrNull(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function bodyOf(attrs: Record<string, unknown>): CommentBody {
  return {
    authorId: stringOrNull(attrs.authorId),
    author: stringOrNull(attrs.author),
    initials: stringOrNull(attrs.initials),
    date: stringOrNull(attrs.date),
    text: stringOrNull(attrs.text) ?? "",
  };
}

function replyBodies(value: unknown): Map<string, CommentBody> {
  const replies = new Map<string, CommentBody>();
  if (!Array.isArray(value)) return replies;
  for (const entry of value) {
    if (typeof entry !== "object" || entry === null) continue;
    const reply: Record<string, unknown> = entry;
    const id = stringOrNull(reply.id);
    if (id === null) continue;
    replies.set(id, bodyOf(reply));
  }
  return replies;
}

/** Every comment thread in the document by id, as much of it as ownership is judged on */
function threadsIn(doc: PMNode): Map<string, CommentThread> {
  const threads = new Map<string, CommentThread>();
  doc.descendants((node) => {
    if (node.type.name !== "commentReference") return true;
    const id = stringOrNull(node.attrs.id);
    if (id === null || threads.has(id)) return true;
    threads.set(id, {
      ...bodyOf(node.attrs),
      replies: replyBodies(node.attrs.replies),
    });
    return true;
  });
  return threads;
}

/**
 * Whether a comment written under this identity may be edited or deleted.
 *
 * A comment carrying no recognised identity belongs to nobody in particular and stays open to
 * everyone, which is what every comment was before identities were recorded and what a comment
 * made in Word is; one carrying an identity is its author's alone.
 */
export function commentOwned(
  rules: ProtectionState,
  authorId: string | null
): boolean {
  if (rules.editableComments === "all") return true;
  return authorId === null || authorId === rules.authorId;
}

/**
 * Where the markers of each comment stand, counted in a document with every marker taken out.
 *
 * Moving a comment onto another sentence leaves the text exactly as it was, so the comparison the
 * markers are taken out of cannot see it. Counting in that same comment-free document is what
 * makes the position tell a move from an edit of what a comment says: adding or taking away
 * another comment moves no marker recorded here, and a rewritten body moves none either.
 */
function commentAnchors(doc: PMNode): Map<string, string> {
  const markers = new Map<string, string[]>();
  const visit = (node: PMNode, start: number): number => {
    if (isCommentNode(node)) {
      const id = stringOrNull(node.attrs.id);
      if (id !== null) {
        const at = markers.get(id) ?? [];
        at.push(`${node.type.name}@${start}`);
        markers.set(id, at);
      }
      return 0;
    }
    if (node.isLeaf) return node.nodeSize;
    let content = 0;
    node.forEach((child) => {
      content += visit(child, start + 1 + content);
    });
    return content + 2;
  };
  // The document itself stands one position before its content, which begins at zero
  visit(doc, -1);
  return new Map(
    Array.from(markers, ([id, at]) => [id, at.join(" ")] as const)
  );
}

/** Whether these two are the same comment written by the same person, whatever either says */
function sameIdentity(was: CommentBody, now: CommentBody): boolean {
  return was.authorId === now.authorId && was.author === now.author;
}

/** Whether what the comment says still reads as it did, the identity aside */
function sameWords(was: CommentBody, now: CommentBody): boolean {
  return (
    was.text === now.text &&
    was.initials === now.initials &&
    was.date === now.date
  );
}

/**
 * Whether every comment and reply that was already there kept the identity it was written under.
 *
 * An identity is not an edit anyone may make, its own author included: a comment rewritten into
 * another author's name would be that author's to edit from then on, and one rewritten into
 * nobody's would be everyone's. Only a comment that appears carries an identity chosen for it.
 */
export function commentIdentitiesKept(before: PMNode, after: PMNode): boolean {
  const now = threadsIn(after);
  for (const [id, thread] of threadsIn(before)) {
    const current = now.get(id);
    if (current === undefined) continue;
    if (!sameIdentity(thread, current)) return false;
    for (const [replyId, reply] of thread.replies) {
      const currentReply = current.replies.get(replyId);
      if (currentReply !== undefined && !sameIdentity(reply, currentReply)) {
        return false;
      }
    }
  }
  return true;
}

/**
 * Whether every comment edited, moved or taken away between the two documents belonged to whoever
 * did it. Adding a comment, replying, and settling a thread as resolved or open belong to
 * everyone.
 *
 * Moving a comment onto other text is an edit of that comment rather than of the body, since the
 * body reads the same afterwards, and it is the owner's to make.
 * A root taken away takes its replies with it whoever wrote them: a thread hangs off its root,
 * and the root's owner may take the thread down.
 */
export function commentEditsOwned(
  before: PMNode,
  after: PMNode,
  rules: ProtectionState
): boolean {
  if (!commentIdentitiesKept(before, after)) return false;
  if (rules.editableComments === "all") return true;
  const now = threadsIn(after);
  const anchoredBefore = commentAnchors(before);
  const anchoredAfter = commentAnchors(after);
  for (const [id, thread] of threadsIn(before)) {
    const current = now.get(id);
    const changed =
      current === undefined ||
      !sameWords(thread, current) ||
      anchoredBefore.get(id) !== anchoredAfter.get(id);
    if (changed && !commentOwned(rules, thread.authorId)) return false;
    if (current === undefined) continue;
    for (const [replyId, reply] of thread.replies) {
      const currentReply = current.replies.get(replyId);
      const replyChanged =
        currentReply === undefined || !sameWords(reply, currentReply);
      if (replyChanged && !commentOwned(rules, reply.authorId)) return false;
    }
  }
  return true;
}

/**
 * Whether the protection lets a document go from `before` to `after`.
 *
 * Ownership is judged only over a change that is nothing but comments. A body edit that sweeps a
 * comment away with the text it was anchored to is a body edit, and the standing to make one is
 * the protection's question; under `comments` no such edit goes through in the first place.
 */
export function protectionAllows(
  before: PMNode,
  after: PMNode,
  rules: ProtectionState
): boolean {
  switch (rules.protection) {
    case "readOnly":
      return false;
    case "comments":
      return (
        changesOnlyComments(before, after) &&
        commentEditsOwned(before, after, rules)
      );
    case "none":
      return (
        !changesOnlyComments(before, after) ||
        commentEditsOwned(before, after, rules)
      );
    default: {
      const unmodelled: never = rules.protection;
      return unmodelled;
    }
  }
}

/**
 * Whether every comment and reply that appeared between the two documents was written under this
 * identity, and every one that was already there kept the identity it was written under. The
 * editor takes an addition from anyone, since it writes the author itself; a server taking a file
 * back does not, since the file could claim any author.
 */
export function commentAdditionsBy(
  before: PMNode,
  after: PMNode,
  authorId: string
): boolean {
  if (!commentIdentitiesKept(before, after)) return false;
  const was = threadsIn(before);
  for (const [id, thread] of threadsIn(after)) {
    const earlier = was.get(id);
    if (earlier === undefined) {
      if (thread.authorId !== authorId) return false;
      for (const reply of thread.replies.values()) {
        if (reply.authorId !== authorId) return false;
      }
      continue;
    }
    for (const [replyId, reply] of thread.replies) {
      if (!earlier.replies.has(replyId) && reply.authorId !== authorId) {
        return false;
      }
    }
  }
  return true;
}
