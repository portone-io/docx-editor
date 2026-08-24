/**
 * Undo and redo, each carrying the pass that takes a replayed lock through the lock guard
 * (`historyReplay` in `schema/locks`).
 *
 * These are the two the keymap, the toolbar and `./commands` use. A raw prosemirror-history command
 * is refused the moment the history reaches a lock, and that refusal leaves everything behind it in
 * the history out of reach.
 */

import { redo as historyRedo, undo as historyUndo } from "prosemirror-history";
import type { Command } from "prosemirror-state";
import { historyReplay } from "../../schema/locks";

/** The same command, with every transaction it dispatches carrying the pass */
function replaying(command: Command): Command {
  return (state, dispatch, view) =>
    command(
      state,
      dispatch && ((tr) => dispatch(tr.setMeta(historyReplay, true))),
      view
    );
}

export const undo = replaying(historyUndo);
export const redo = replaying(historyRedo);
