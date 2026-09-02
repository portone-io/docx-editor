/**
 * Building a transaction and running it are kept apart.
 * Following the ProseMirror convention, calling a command without `dispatch` only
 * reports whether it can run right now, and that answer takes in the lock guard
 * (`schema/locks`), so no caller has to ask about locks separately.
 */

import type { EditorState, Transaction } from "prosemirror-state";
import {
  deleteTable as pmDeleteTable,
  isInTable as pmIsInTable,
} from "prosemirror-tables";
import { transactionAllowed } from "../schema/locks";
import {
  buildAddColumnAfterTransaction,
  buildAddColumnBeforeTransaction,
  buildDeleteColumnTransaction,
} from "./columns";
import { buildMergeCellsTransaction, buildSplitCellTransaction } from "./merge";
import {
  buildAddRowAfterTransaction,
  buildAddRowBeforeTransaction,
  buildDeleteRowTransaction,
} from "./rows";

export type TableCommand = (
  state: EditorState,
  dispatch?: (tr: Transaction) => void
) => boolean;

/** Whether the cursor sits inside a table. Used to enable and disable the table buttons */
export function isInTable(state: EditorState): boolean {
  return pmIsInTable(state);
}

/**
 * A table command out of the transaction it builds, refused where the lock guard would turn that
 * transaction down (`schema/locks`).
 *
 * A structural edit is refused whole rather than trimmed: half a row cannot be deleted, and half a
 * block of cells cannot be merged into one. That is the opposite of a paragraph or character edit,
 * which leaves the locked stretches out and applies to the rest (`editor/paragraphEdits`).
 *
 * The answer is the same whether or not `dispatch` was passed. The transaction is built before
 * either way, so asking the guard costs nothing more, and a command reporting one thing to a button
 * and doing another would be worse than the button being wrong.
 */
function toCommand(
  build: (state: EditorState) => Transaction | null
): TableCommand {
  return (state, dispatch) => {
    const tr = build(state);
    if (!tr || !transactionAllowed(tr, state)) return false;
    dispatch?.(tr);
    return true;
  };
}

export const addRowBefore: TableCommand = toCommand(
  buildAddRowBeforeTransaction
);
export const addRowAfter: TableCommand = toCommand(buildAddRowAfterTransaction);
export const deleteRow: TableCommand = toCommand(buildDeleteRowTransaction);

export const addColumnBefore: TableCommand = toCommand(
  buildAddColumnBeforeTransaction
);
export const addColumnAfter: TableCommand = toCommand(
  buildAddColumnAfterTransaction
);
export const deleteColumn: TableCommand = toCommand(
  buildDeleteColumnTransaction
);

function buildDeleteTableTransaction(state: EditorState): Transaction | null {
  if (!pmIsInTable(state)) return null;
  const captured: Transaction[] = [];
  pmDeleteTable(state, (tr) => captured.push(tr));
  return captured[0] ?? null;
}

export const deleteTable: TableCommand = toCommand(buildDeleteTableTransaction);

export const mergeCells: TableCommand = toCommand(buildMergeCellsTransaction);
export const splitCell: TableCommand = toCommand(buildSplitCellTransaction);
