/**
 * A table only goes between body blocks, never inside another table. A table inside a cell is not
 * made editable even when a document is opened (nested tables are only preserved), so one created
 * here would become an untouchable blob once the document is saved and reopened.
 */

import { Fragment, type ResolvedPos } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { createTableNode, isTableSide } from "../docx/tableTemplate";
import { editsShut } from "../schema/protectionState";
import { documentGeometry } from "./documentStyles";

/** Whether this position is inside a table */
function isInTable($pos: ResolvedPos): boolean {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    if ($pos.node(depth).type.spec.tableRole === "table") return true;
  }
  return false;
}

/**
 * Where the new table goes.
 * Right after the body block the caret sits in, so the text being written is not split in two
 * by the table. Null when the table cannot be inserted here, a protection shutting the body included.
 */
function insertPosition(state: EditorState): number | null {
  const $from = state.selection.$from;
  if (editsShut(state) || $from.depth === 0 || isInTable($from)) return null;
  return $from.after(1);
}

/** Whether a table can be inserted at the current position. Used to enable and disable the toolbar button */
export function canInsertTable(state: EditorState): boolean {
  return insertPosition(state) !== null;
}

/** How many rows and columns the new table has */
export interface TableSize {
  rows: number;
  columns: number;
}

/**
 * Inserts a table of empty cells and moves the caret into the first cell.
 * The table is as wide as one line of body text on the paper the open document names
 */
export function insertTable({ rows, columns }: TableSize): Command {
  return (state, dispatch) => {
    const at = insertPosition(state);
    if (at === null || !isTableSide(rows) || !isTableSide(columns))
      return false;
    if (dispatch) {
      const tr = state.tr.insert(
        at,
        Fragment.fromArray([
          createTableNode(rows, columns, documentGeometry(state)),
          state.schema.nodes.paragraph.create(),
        ])
      );
      const selection = TextSelection.near(tr.doc.resolve(at + 1));
      dispatch(tr.setSelection(selection).scrollIntoView());
    }
    return true;
  };
}
