/** Edits Word comments through their range markers and reference nodes. */

import type { Mark, Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { commentParaId } from "../../../docx/comments";
import { docxSchema } from "../../../schema";
import { transactionAllowed } from "../../../schema/locks";
import {
  reservedCommentIds,
  reservedCommentParaIds,
} from "../../plugins/commentReservations";
import type { NewComment } from "./model";
import { documentComments, repliesAttr, stringAttr } from "./reading";

function validCommentSelection(state: EditorState): boolean {
  const { selection } = state;
  return (
    selection instanceof TextSelection &&
    !selection.empty &&
    selection.$from.sameParent(selection.$to) &&
    selection.$from.parent.type.name === "paragraph"
  );
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
  return (adjacent?.marks ?? resolved.marks()).filter(
    (mark) => mark.type.name === "sdt" || mark.type.name === "link"
  );
}

function addCommentTransaction(
  state: EditorState,
  comment: NewComment
): Transaction | null {
  if (!validCommentSelection(state) || comment.text.trim().length === 0) {
    return null;
  }
  const id = nextCommentId(state);
  const date = comment.date ?? new Date().toISOString();
  const paraId = nextCommentParaId(state, `comment-${id}-${date}`);
  const { from, to } = state.selection;
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
      initials: comment.initials ?? null,
      date,
      text: comment.text,
      commentXml: null,
      imported: false,
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
  const transaction = state.tr
    .insert(to, end)
    .insert(to + 1, reference)
    .insert(from, start);
  return transactionAllowed(transaction, state) ? transaction : null;
}

/** Whether a non-empty selection in one paragraph can receive a comment. */
export function canAddComment(state: EditorState): boolean {
  return (
    addCommentTransaction(state, {
      text: "comment",
      author: "Author",
      date: "1970-01-01T00:00:00.000Z",
    }) !== null
  );
}

/** Adds a plain-text comment to the current text selection. */
export function addComment(comment: NewComment): Command {
  return (state, dispatch) => {
    const transaction = addCommentTransaction(state, comment);
    if (!transaction) return false;
    dispatch?.(transaction);
    return true;
  };
}

/** Replaces the plain-text body of one comment, retaining its author and anchor. */
export function updateComment(id: string, text: string): Command {
  return (state, dispatch) => {
    if (text.trim().length === 0) return false;
    let transaction = state.tr;
    let changed = false;
    state.doc.descendants((node, pos) => {
      if (
        node.type.name === "commentReference" &&
        stringAttr(node.attrs.id) === id &&
        stringAttr(node.attrs.text) !== text
      ) {
        transaction = transaction.setNodeMarkup(pos, null, {
          ...node.attrs,
          text,
          commentXml: null,
          imported: false,
        });
        changed = true;
      }
      return true;
    });
    if (!changed || !transactionAllowed(transaction, state)) return false;
    dispatch?.(transaction);
    return true;
  };
}

function updateReference(
  id: string,
  change: (node: PMNode) => Record<string, unknown> | null
): Command {
  return (state, dispatch) => {
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
    if (!changed || !transactionAllowed(transaction, state)) return false;
    dispatch?.(transaction);
    return true;
  };
}

/** Marks a comment thread resolved or open without deleting it. */
export function setCommentResolved(id: string, resolved: boolean): Command {
  return (state, dispatch) =>
    updateReference(id, (node) => {
      if (node.attrs.resolved === resolved) return null;
      return {
        ...node.attrs,
        paraId:
          node.attrs.threadImported === true && node.attrs.extensionXml === null
            ? nextCommentParaId(state, `comment-${id}`)
            : node.attrs.paraId,
        resolved,
        extensionXml: null,
        threadImported: false,
        imported:
          node.attrs.commentXml !== null && node.attrs.extensionXml === null
            ? false
            : node.attrs.imported,
      };
    })(state, dispatch);
}

/** Adds a plain-text reply and reopens the thread when it was resolved. */
export function addCommentReply(id: string, reply: NewComment): Command {
  return (state, dispatch) => {
    if (reply.text.trim().length === 0) return false;
    const replyId = nextCommentId(state);
    const date = reply.date ?? new Date().toISOString();
    return updateReference(id, (node) => {
      const parentParaId =
        node.attrs.threadImported === true && node.attrs.extensionXml === null
          ? nextCommentParaId(state, `comment-${id}`)
          : (stringAttr(node.attrs.paraId) ??
            nextCommentParaId(state, `comment-${id}`));
      const paraId = nextCommentParaId(state, `comment-${replyId}-${date}`, [
        parentParaId,
      ]);
      return {
        ...node.attrs,
        paraId: parentParaId,
        resolved: false,
        extensionXml: null,
        threadImported: false,
        imported:
          node.attrs.commentXml !== null && node.attrs.extensionXml === null
            ? false
            : node.attrs.imported,
        replies: [
          ...repliesAttr(node.attrs.replies),
          {
            id: replyId,
            author: reply.author,
            initials: reply.initials ?? null,
            date,
            text: reply.text,
            commentXml: null,
            imported: false,
            paraId,
            parentParaId,
            extensionXml: null,
          },
        ],
      };
    })(state, dispatch);
  };
}

/** Replaces the plain-text body of one reply. */
export function updateCommentReply(
  commentId: string,
  replyId: string,
  text: string
): Command {
  return updateReference(commentId, (node) => {
    if (text.trim().length === 0) return null;
    const replies = repliesAttr(node.attrs.replies);
    const index = replies.findIndex((reply) => reply.id === replyId);
    if (index < 0 || replies[index].text === text) return null;
    return {
      ...node.attrs,
      replies: replies.map((reply, at) =>
        at === index
          ? { ...reply, text, commentXml: null, imported: false }
          : reply
      ),
    };
  });
}

/** Removes one reply while retaining its root comment and anchor. */
export function removeCommentReply(
  commentId: string,
  replyId: string
): Command {
  return updateReference(commentId, (node) => {
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
    return {
      ...node.attrs,
      replies: replies.filter((reply) => !removedIds.has(reply.id)),
      threadImported: false,
    };
  });
}

/** Removes a comment's range markers, reference and Comments-part entry. */
export function removeComment(id: string): Command {
  return (state, dispatch) => {
    const positions: Array<{ pos: number; size: number }> = [];
    state.doc.descendants((node, pos) => {
      if (
        (node.type.name === "commentStart" ||
          node.type.name === "commentEnd" ||
          node.type.name === "commentReference") &&
        stringAttr(node.attrs.id) === id
      ) {
        positions.push({ pos, size: node.nodeSize });
      }
      return true;
    });
    if (positions.length === 0) return false;
    const transaction = state.tr;
    for (const marker of positions.sort((a, b) => b.pos - a.pos)) {
      transaction.delete(marker.pos, marker.pos + marker.size);
    }
    if (!transactionAllowed(transaction, state)) return false;
    dispatch?.(transaction);
    return true;
  };
}

/** Selects the text anchored by a comment, or places the caret at a point comment. */
export function selectComment(id: string): Command {
  return (state, dispatch) => {
    const comment = documentComments(state).find((entry) => entry.id === id);
    if (!comment) return false;
    dispatch?.(
      state.tr.setSelection(
        TextSelection.create(state.doc, comment.from, comment.to)
      )
    );
    return true;
  };
}
