/** Runs editor commands and restores focus for continued typing. */

import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

export type RunCommand = (command: Command) => void;

function apply(view: EditorView, command: Command): void {
  command(view.state, (tr) => view.dispatch(tr), view);
}

/**
 * Runs commands on `view` and leaves the focus on `focus`, which is `view` itself unless the
 * command acts somewhere other than where the caret stands.
 *
 * Undo and redo are the one case: they run on the document while the caret is in a story, and
 * handing the focus to the document is what closes an open story (`DocxEditor`).
 */
export function commandRunner(
  view: EditorView,
  focus: EditorView = view
): RunCommand {
  return (command) => {
    apply(view, command);
    focus.focus();
  };
}
