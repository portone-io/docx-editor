/** Runs editor commands and restores focus for continued typing. */

import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

export type RunCommand = (command: Command) => void;

function apply(view: EditorView, command: Command): void {
  command(view.state, (tr) => view.dispatch(tr), view);
}

export function commandRunner(view: EditorView): RunCommand {
  return (command) => {
    apply(view, command);
    view.focus();
  };
}
