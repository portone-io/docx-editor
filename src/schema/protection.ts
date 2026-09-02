/**
 * The protection an editor runs under: what the document as a whole may receive, over and above
 * the locks its own controls carry (`./locks`).
 *
 * The levels are named after OOXML `ST_DocProtect`, the enumeration `w:documentProtection/@w:edit`
 * takes (ECMA-376 Part 1 §17.15.1.29): `readOnly` takes no edit at all, `comments` takes comments
 * and nothing else, `none` takes everything. `trackedChanges` and `forms`, the two values of the
 * enumeration not modelled yet, would be added here.
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

interface CommentBody {
  authorId: string | null;
  text: string;
}

interface CommentThread extends CommentBody {
  replies: ReadonlyMap<string, CommentBody>;
}

function stringOrNull(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function replyBodies(value: unknown): Map<string, CommentBody> {
  const replies = new Map<string, CommentBody>();
  if (!Array.isArray(value)) return replies;
  for (const entry of value) {
    if (typeof entry !== "object" || entry === null) continue;
    const reply: Record<string, unknown> = entry;
    const id = stringOrNull(reply.id);
    if (id === null) continue;
    replies.set(id, {
      authorId: stringOrNull(reply.authorId),
      text: stringOrNull(reply.text) ?? "",
    });
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
      authorId: stringOrNull(node.attrs.authorId),
      text: stringOrNull(node.attrs.text) ?? "",
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

/** Whether a body that was there before may have become what it is now: unchanged, or changed by its owner */
function bodyOwned(
  rules: ProtectionState,
  was: CommentBody,
  now: CommentBody | undefined
): boolean {
  if (now !== undefined && now.text === was.text) return true;
  return commentOwned(rules, was.authorId);
}

/**
 * Whether every comment edited or taken away between the two documents belonged to whoever did
 * it. Adding a comment, replying, and settling a thread as resolved or open belong to everyone.
 *
 * A root taken away takes its replies with it whoever wrote them: a thread hangs off its root,
 * and the root's owner may take the thread down.
 */
export function commentEditsOwned(
  before: PMNode,
  after: PMNode,
  rules: ProtectionState
): boolean {
  if (rules.editableComments === "all") return true;
  const now = threadsIn(after);
  for (const [id, thread] of threadsIn(before)) {
    const current = now.get(id);
    if (!bodyOwned(rules, thread, current)) return false;
    if (current === undefined) continue;
    for (const [replyId, reply] of thread.replies) {
      if (!bodyOwned(rules, reply, current.replies.get(replyId))) return false;
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
  if (rules.protection === "readOnly") return false;
  const onlyComments = changesOnlyComments(before, after);
  if (rules.protection === "comments" && !onlyComments) return false;
  return !onlyComments || commentEditsOwned(before, after, rules);
}

/**
 * Whether every comment and reply that appeared between the two documents was written under this
 * identity. The editor takes an addition from anyone, since it writes the author itself; a server
 * taking a file back does not, since the file could claim any author.
 */
export function commentAdditionsBy(
  before: PMNode,
  after: PMNode,
  authorId: string
): boolean {
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
