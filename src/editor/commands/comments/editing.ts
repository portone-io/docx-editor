/** Edits Word comments through their range markers and reference nodes. */

import type { Mark, Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { commentParaId } from "../../../docx/comments";
import {
  sameStory,
  setStory,
  storyFromText,
  storyKey,
  storyOf,
  storyText,
  withoutStories,
} from "../../../docx/story";
import { docxSchema } from "../../../schema";
import { guardedCommand } from "../../../schema/guards";
import { isWrapperType } from "../../../schema/wrappers";
import {
  reservedCommentIds,
  reservedCommentParaIds,
} from "../../editorDocument";
import {
  type CommentRange,
  type NewComment,
  repliesAttr,
  stringAttr,
} from "./model";
import { commentById } from "./reading";

/**
 * The stretch as a comment anchor may stand over it, or null where one may not: a comment marks a
 * run of text inside one paragraph, so an empty stretch, one reaching past the paragraph it starts
 * in, or one naming positions the document does not have takes none.
 */
function anchorableRange(
  doc: PMNode,
  { from, to }: CommentRange
): CommentRange | null {
  if (
    !Number.isInteger(from) ||
    !Number.isInteger(to) ||
    from >= to ||
    from < 0 ||
    to > doc.content.size
  )
    return null;
  const $from = doc.resolve(from);
  if (!$from.sameParent(doc.resolve(to))) return null;
  if ($from.parent.type.name !== "paragraph") return null;
  return { from, to };
}

/**
 * The stretch this comment is written for: the one named, or the selected text when none is.
 *
 * A selection that is not a run of text - a picture, a block of table cells - is no stretch to
 * comment on, while a range named outright is the caller's own reading of the document and is
 * judged on where it lands alone.
 */
function commentedRange(
  state: EditorState,
  at: CommentRange | undefined
): CommentRange | null {
  if (at) return anchorableRange(state.doc, at);
  const { selection } = state;
  if (!(selection instanceof TextSelection)) return null;
  return anchorableRange(state.doc, selection);
}

function nextCommentId(state: EditorState): string {
  const taken = new Set<string>();
  const reserve = (id: string) => {
    if (!/^\d+$/.test(id)) return;
    try {
      taken.add(BigInt(id).toString());
    } catch {
      return;
    }
  };
  for (const id of reservedCommentIds(state)) {
    reserve(id);
  }
  state.doc.descendants((node) => {
    if (
      node.type.name === "commentStart" ||
      node.type.name === "commentEnd" ||
      node.type.name === "commentReference"
    ) {
      const id = stringAttr(node.attrs.id);
      if (id !== null) reserve(id);
    }
    if (node.type.name === "commentReference") {
      for (const reply of repliesAttr(node.attrs.replies)) {
        reserve(reply.id);
      }
    }
    return true;
  });
  let max = -1n;
  for (const id of taken) {
    const value = BigInt(id);
    if (value > max) max = value;
  }
  return (max + 1n).toString();
}

function nextCommentParaId(
  state: EditorState,
  seed: string,
  additional: Iterable<string> = []
): string {
  const taken = new Set(
    Array.from(reservedCommentParaIds(state), (id) => id.toUpperCase())
  );
  state.doc.descendants((node) => {
    if (node.type.name !== "commentReference") return true;
    const paraId = stringAttr(node.attrs.paraId);
    if (paraId !== null) taken.add(paraId.toUpperCase());
    for (const reply of repliesAttr(node.attrs.replies)) {
      taken.add(reply.paraId.toUpperCase());
    }
    return true;
  });
  for (const paraId of additional) taken.add(paraId.toUpperCase());
  let candidate = Number.parseInt(commentParaId(seed), 16);
  while (taken.has(candidate.toString(16).toUpperCase().padStart(8, "0"))) {
    candidate = candidate === 0x7fffffff ? 1 : candidate + 1;
  }
  return candidate.toString(16).toUpperCase().padStart(8, "0");
}

function wrapperMarksAt(
  state: EditorState,
  pos: number,
  side: "before" | "after"
): readonly Mark[] {
  const resolved = state.doc.resolve(pos);
  const adjacent = side === "before" ? resolved.nodeBefore : resolved.nodeAfter;
  return (adjacent?.marks ?? resolved.marks()).filter((mark) =>
    isWrapperType(mark.type)
  );
}

function addCommentTransaction(
  state: EditorState,
  comment: NewComment,
  at: CommentRange | undefined
): Transaction | null {
  const range = commentedRange(state, at);
  if (range === null || comment.text.trim().length === 0) return null;
  const id = nextCommentId(state);
  const date = comment.date ?? new Date().toISOString();
  const paraId = nextCommentParaId(state, `comment-${id}-${date}`);
  const { from, to } = range;
  const startMarks = wrapperMarksAt(state, from, "after");
  const endMarks = wrapperMarksAt(state, to, "before");
  const end = docxSchema.nodes.commentEnd.create(
    { id, xml: null },
    null,
    endMarks
  );
  const reference = docxSchema.nodes.commentReference.create(
    {
      id,
      referenceXml: null,
      author: comment.author,
      authorId: comment.authorId ?? null,
      initials: comment.initials ?? null,
      date,
      paraId,
      resolved: false,
      extensionXml: null,
      threadImported: true,
      replies: [],
    },
    null,
    endMarks
  );
  const start = docxSchema.nodes.commentStart.create(
    { id, xml: null },
    null,
    startMarks
  );
  return setStory(
    state.tr
      .insert(to, end)
      .insert(to + 1, reference)
      .insert(from, start),
    storyKey("comment", id),
    storyFromText(comment.text)
  );
}

/**
 * Adds a plain-text comment to a stretch of text: the one given, or the current text selection.
 *
 * A composer that was opened over one stretch and submitted later names it, so that what the
 * comment marks is the text it was written about rather than wherever the caret has since gone.
 */
export function addComment(comment: NewComment, at?: CommentRange): Command {
  return guardedCommand((state) => addCommentTransaction(state, comment, at));
}

/**
 * Whether that stretch of text, or the current selection, can receive a comment.
 * This is the command itself asked without a dispatch, over a comment standing in for the one the
 * user would write, so the button and the click cannot answer differently.
 */
export function canAddComment(state: EditorState, at?: CommentRange): boolean {
  return addComment(
    {
      text: "comment",
      author: "Author",
      date: "1970-01-01T00:00:00.000Z",
    },
    at
  )(state);
}

/** Where the first reference to this comment stands, and null where the document holds none */
function referenceAt(state: EditorState, id: string): number | null {
  let at: number | null = null;
  state.doc.descendants((node, pos) => {
    if (at !== null) return false;
    if (
      node.type.name === "commentReference" &&
      stringAttr(node.attrs.id) === id
    ) {
      at = pos;
    }
    return at === null;
  });
  return at;
}

/**
 * Writes one body, comment or reply, as the story the document holds it in.
 *
 * The transaction also rewrites the reference it belongs to, attr for attr. What the comment says
 * now stands on the document node rather than on that node, and a transaction reaching neither the
 * marker nor the reference is not a comment edit as far as the guards can tell: a comment
 * protection would turn it down for touching no comment, and a lock over the text it is anchored
 * in would let it through.
 *
 * null for a body nobody would take: one the document refers to no comment for, one saying
 * nothing, and one already saying what it says.
 */
function bodyTransaction(
  state: EditorState,
  referenceId: string,
  bodyId: string,
  body: PMNode
): Transaction | null {
  if (body.type !== docxSchema.nodes.doc) return null;
  if (storyText(body).trim().length === 0) return null;
  const at = referenceAt(state, referenceId);
  const reference = at === null ? null : state.doc.nodeAt(at);
  if (at === null || reference === null) return null;
  const key = storyKey("comment", bodyId);
  // What a block would be written back as, rather than `Node.eq`: opening a file works the display
  // values out again (`schema/attrRoles`), and a body rebuilt over that would lose the markup the
  // writer does not model
  if (sameStory(storyOf(state.doc, key), body)) return null;
  return setStory(state.tr, key, body).setNodeMarkup(
    at,
    null,
    reference.attrs,
    reference.marks
  );
}

/**
 * The same for a body given as plain text, which is a no-op where the body already reads as that
 * text.
 *
 * Text says what a comment should read as and nothing about how it is written, so a body that
 * already reads this way is left as it stands, its run formatting and its second paragraph with
 * it. Writing how a body is put together is `setCommentBody`'s.
 */
function textBodyTransaction(
  state: EditorState,
  referenceId: string,
  bodyId: string,
  text: string
): Transaction | null {
  const held = storyOf(state.doc, storyKey("comment", bodyId));
  if (held !== null && storyText(held) === text) return null;
  return bodyTransaction(state, referenceId, bodyId, storyFromText(text));
}

/** Replaces the plain-text body of one comment, retaining its author and anchor. */
export function updateComment(id: string, text: string): Command {
  return guardedCommand((state) => textBodyTransaction(state, id, id, text));
}

/**
 * Replaces the body of one comment with a formatted one, retaining its author and anchor.
 *
 * The body is a document of `docxSchema`, the same schema the page is edited in, so a paragraph
 * style, a bold run and a second paragraph all survive being written, exported and read back.
 * `updateComment` is the same command for a body that is only text.
 */
export function setCommentBody(id: string, body: PMNode): Command {
  return guardedCommand((state) => bodyTransaction(state, id, id, body));
}

/**
 * Rewrites the attrs of one comment reference, and whatever else the same edit writes beside them:
 * a reply's own body is a story of its own, and it lands in the transaction that adds the reply so
 * that one undo takes both back.
 */
function updateReference(
  id: string,
  change: (node: PMNode) => Record<string, unknown> | null,
  alongside: (tr: Transaction) => Transaction = (tr) => tr
): Command {
  return guardedCommand((state) => {
    let transaction = state.tr;
    let changed = false;
    state.doc.descendants((node, pos) => {
      if (
        node.type.name !== "commentReference" ||
        stringAttr(node.attrs.id) !== id
      ) {
        return true;
      }
      const attrs = change(node);
      if (attrs !== null) {
        transaction = transaction.setNodeMarkup(pos, null, attrs);
        changed = true;
      }
      return false;
    });
    return changed ? alongside(transaction) : null;
  });
}

/** Marks a comment thread resolved or open without deleting it. */
export function setCommentResolved(id: string, resolved: boolean): Command {
  return (state, dispatch) =>
    updateReference(id, (node) => {
      if (node.attrs.resolved === resolved) return null;
      // The key the thread state hangs off is the one the comment already has, whether it arrived
      // with it or was given one on the way in. The entry keeps whatever it says, and the writer
      // puts the key on it (`docx/comments/grammar`)
      return {
        ...node.attrs,
        resolved,
        extensionXml: null,
        threadImported: false,
      };
    })(state, dispatch);
}

/** Adds a plain-text reply and reopens the thread when it was resolved. */
export function addCommentReply(id: string, reply: NewComment): Command {
  return (state, dispatch) => {
    if (reply.text.trim().length === 0) return false;
    const replyId = nextCommentId(state);
    const date = reply.date ?? new Date().toISOString();
    return updateReference(
      id,
      (node) => {
        // The key a reply hangs off is the one the comment already has, whether it arrived with
        // it or was given one on the way in. Minting a second would re-point the thread
        const parentParaId =
          stringAttr(node.attrs.paraId) ??
          nextCommentParaId(state, `comment-${id}`);
        const paraId = nextCommentParaId(state, `comment-${replyId}-${date}`, [
          parentParaId,
        ]);
        return {
          ...node.attrs,
          paraId: parentParaId,
          resolved: false,
          extensionXml: null,
          threadImported: false,
          replies: [
            ...repliesAttr(node.attrs.replies),
            {
              id: replyId,
              author: reply.author,
              authorId: reply.authorId ?? null,
              initials: reply.initials ?? null,
              date,
              paraId,
              parentParaId,
              extensionXml: null,
            },
          ],
        };
      },
      (tr) =>
        setStory(tr, storyKey("comment", replyId), storyFromText(reply.text))
    )(state, dispatch);
  };
}

/** Replaces the plain-text body of one reply. */
export function updateCommentReply(
  commentId: string,
  replyId: string,
  text: string
): Command {
  return guardedCommand((state) => {
    const holdsReply = repliesOf(state, commentId).some(
      (reply) => reply.id === replyId
    );
    if (!holdsReply) return null;
    return textBodyTransaction(state, commentId, replyId, text);
  });
}

/** The replies one comment reference carries, empty for an id the document refers to no comment for */
function repliesOf(
  state: EditorState,
  commentId: string
): readonly { id: string }[] {
  let replies: readonly { id: string }[] = [];
  state.doc.descendants((node) => {
    if (
      node.type.name !== "commentReference" ||
      stringAttr(node.attrs.id) !== commentId
    ) {
      return true;
    }
    replies = repliesAttr(node.attrs.replies);
    return false;
  });
  return replies;
}

/** The same transaction with the bodies of these comments taken off the document node */
function dropBodies(ids: Iterable<string>): (tr: Transaction) => Transaction {
  return (tr) =>
    withoutStories(
      tr,
      Array.from(ids, (id) => storyKey("comment", id))
    );
}

/** Removes one reply while retaining its root comment and anchor. */
export function removeCommentReply(
  commentId: string,
  replyId: string
): Command {
  return (state, dispatch) => {
    // What comes down with the reply is worked out against the document this run is made over. A
    // command is asked whether it applies and then asked to run, so a set the command itself held
    // would carry one document's answer into the next
    const removedReplyIds = new Set<string>();
    return updateReference(
      commentId,
      (node) => {
        const replies = repliesAttr(node.attrs.replies);
        if (!replies.some((reply) => reply.id === replyId)) return null;
        const removedIds = new Set([replyId]);
        const removedParaIds = new Set(
          replies
            .filter((reply) => removedIds.has(reply.id))
            .map((reply) => reply.paraId)
        );
        let changed = true;
        while (changed) {
          changed = false;
          for (const reply of replies) {
            if (
              !removedIds.has(reply.id) &&
              removedParaIds.has(reply.parentParaId)
            ) {
              removedIds.add(reply.id);
              removedParaIds.add(reply.paraId);
              changed = true;
            }
          }
        }
        for (const id of removedIds) removedReplyIds.add(id);
        return {
          ...node.attrs,
          replies: replies.filter((reply) => !removedIds.has(reply.id)),
          threadImported: false,
        };
      },
      dropBodies(removedReplyIds)
    )(state, dispatch);
  };
}

/**
 * Removes a comment's range markers, reference and Comments-part entry.
 *
 * What the comment and its replies said goes with them: a body is a story on the document node,
 * and one left behind for a comment nothing refers to would ride every later comparison of the
 * document without standing for anything.
 */
export function removeComment(id: string): Command {
  return guardedCommand((state) => {
    const positions: Array<{ pos: number; size: number }> = [];
    const bodies = new Set([id]);
    state.doc.descendants((node, pos) => {
      if (
        (node.type.name === "commentStart" ||
          node.type.name === "commentEnd" ||
          node.type.name === "commentReference") &&
        stringAttr(node.attrs.id) === id
      ) {
        positions.push({ pos, size: node.nodeSize });
      }
      if (
        node.type.name === "commentReference" &&
        stringAttr(node.attrs.id) === id
      ) {
        for (const reply of repliesAttr(node.attrs.replies))
          bodies.add(reply.id);
      }
      return true;
    });
    if (positions.length === 0) return null;
    const transaction = state.tr;
    for (const marker of positions.sort((a, b) => b.pos - a.pos)) {
      transaction.delete(marker.pos, marker.pos + marker.size);
    }
    return dropBodies(bodies)(transaction);
  });
}

/** Selects the text anchored by a comment, or places the caret at a point comment. */
export function selectComment(id: string): Command {
  return (state, dispatch) => {
    const comment = commentById(state, id);
    if (!comment) return false;
    dispatch?.(
      state.tr.setSelection(
        TextSelection.create(state.doc, comment.from, comment.to)
      )
    );
    return true;
  };
}
