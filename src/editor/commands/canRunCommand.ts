/**
 * Dry-runs external commands and checks every produced transaction against document locks without
 * dispatching them or triggering the locked-content plugin's side effects.
 */

import type { Command, EditorState, Transaction } from "prosemirror-state";
import { transactionAllowed } from "../../schema/locks";

export function canRunCommand(command: Command, state: EditorState): boolean {
  const built: Transaction[] = [];
  if (!command(state, (tr) => built.push(tr))) return false;
  return built.every((tr) => transactionAllowed(tr, state.doc));
}
