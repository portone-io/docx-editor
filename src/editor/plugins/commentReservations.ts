/** Keeps identifiers from the opened Comments parts reserved for the lifetime of the editor state. */

import { type EditorState, Plugin, PluginKey } from "prosemirror-state";

interface CommentReservations {
  ids: ReadonlySet<string>;
  paraIds: ReadonlySet<string>;
}

const commentReservationsKey = new PluginKey<CommentReservations>(
  "docxCommentReservations"
);

export function commentReservations(
  ids: Iterable<string>,
  paraIds: Iterable<string>
): Plugin {
  const reserved = { ids: new Set(ids), paraIds: new Set(paraIds) };
  return new Plugin({
    key: commentReservationsKey,
    state: {
      init: () => reserved,
      apply: (_transaction, value) => value,
    },
  });
}

export function reservedCommentIds(state: EditorState): ReadonlySet<string> {
  return commentReservationsKey.getState(state)?.ids ?? new Set();
}

export function reservedCommentParaIds(
  state: EditorState
): ReadonlySet<string> {
  return commentReservationsKey.getState(state)?.paraIds ?? new Set();
}
