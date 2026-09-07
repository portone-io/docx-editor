/**
 * Building a transaction and running it are kept apart.
 * Following the ProseMirror convention, calling a command without `dispatch` only
 * reports whether it can run right now, and that answer takes in every guard an edit
 * is judged by (`schema/guards`), so no caller has to ask about them separately.
 */

import type { EditorState, Transaction } from "prosemirror-state";
import {
  deleteTable as pmDeleteTable,
  isInTable as pmIsInTable,
} from "prosemirror-tables";
import { guardedCommand } from "../schema/guards";
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

export const addRowBefore: TableCommand = guardedCommand(
  buildAddRowBeforeTransaction
);
export const addRowAfter: TableCommand = guardedCommand(
  buildAddRowAfterTransaction
);
export const deleteRow: TableCommand = guardedCommand(
  buildDeleteRowTransaction
);

export const addColumnBefore: TableCommand = guardedCommand(
  buildAddColumnBeforeTransaction
);
export const addColumnAfter: TableCommand = guardedCommand(
  buildAddColumnAfterTransaction
);
export const deleteColumn: TableCommand = guardedCommand(
  buildDeleteColumnTransaction
);

function buildDeleteTableTransaction(state: EditorState): Transaction | null {
  if (!pmIsInTable(state)) return null;
  const captured: Transaction[] = [];
  pmDeleteTable(state, (tr) => captured.push(tr));
  return captured[0] ?? null;
}

export const deleteTable: TableCommand = guardedCommand(
  buildDeleteTableTransaction
);

export const mergeCells: TableCommand = guardedCommand(
  buildMergeCellsTransaction
);
export const splitCell: TableCommand = guardedCommand(
  buildSplitCellTransaction
);
