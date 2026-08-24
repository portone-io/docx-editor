/**
 * In docx a vertically merged cell exists as a single cell only in the row where the
 * merge starts; in the rows below there is no cell at that position at all (how many
 * rows the starting cell covers is counted by rowspan).
 * So the number of cells can differ from row to row, and counting by "the nth cell"
 * goes wrong. Everything here counts by the table's grid coordinates (TableMap).
 *
 * Deletion is left entirely to `prosemirror-tables`' `deleteRow`, which gets shrinking
 * rowspans right as well.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  type EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import {
  deleteRow as pmDeleteRow,
  isInTable as pmIsInTable,
  selectedRect,
  tableNodeTypes,
} from "prosemirror-tables";
import { spanCount } from "../model/format";
import {
  colwidthOf,
  inheritCellAttrs,
  inheritRowAttrs,
  type TableGridMap,
  type TableRect,
} from "./format";

/**
 * If this position points at the same cell as the row above, it is a position covered
 * by a merged cell reaching down from above
 */
function isCoveredFromAbove(
  map: TableGridMap,
  row: number,
  index: number
): boolean {
  return (
    row > 0 && row < map.height && map.map[index] === map.map[index - map.width]
  );
}

/**
 * Grows by one the height of a merged cell that crosses the new row.
 * Returns how many columns that cell occupies (that many have to be skipped).
 */
function growRowspan(tr: Transaction, rect: TableRect, index: number): number {
  const { map, table, tableStart } = rect;
  const pos = map.map[index];
  const cell = table.nodeAt(pos);
  if (!cell) return 1;

  tr.setNodeMarkup(tableStart + pos, null, {
    ...cell.attrs,
    rowspan: spanCount(cell.attrs.rowspan) + 1,
  });
  return spanCount(cell.attrs.colspan);
}

/**
 * Creates an empty cell that inherits only the formatting of the cell at the same
 * position in the reference row.
 * A horizontal merge is reproduced as one cell carrying a span, not as several
 * individual cells.
 */
function inheritedCell(
  rect: TableRect,
  index: number,
  referenceRowOffset: number
): { node: PMNode | null; colspan: number } {
  const { map, table } = rect;
  const referencePos = map.map[index + referenceRowOffset * map.width];
  const referenceCell =
    referencePos != null ? table.nodeAt(referencePos) : null;
  const type = referenceCell
    ? referenceCell.type
    : tableNodeTypes(table.type.schema).cell;
  // The new cell occupies the same grid columns as the reference cell, so its span
  // stays the same and its height is a single row.
  const colspan = referenceCell ? spanCount(referenceCell.attrs.colspan) : 1;
  const attrs = referenceCell
    ? inheritCellAttrs(referenceCell, {
        colspan,
        rowspan: 1,
        colwidth: colwidthOf(referenceCell),
      })
    : null;
  return { node: type.createAndFill(attrs), colspan };
}

/**
 * Inserts a new row at grid row `row` and returns that row's position.
 * If `row` equals the table's row count, the row is appended at the bottom.
 *
 * The actual insertion happens only once, at the very end, so every position computed
 * before it stays valid throughout. Slipping another insertion in between breaks that
 * assumption.
 *
 * When a horizontal and a vertical merge overlap in one cell, the cell to its right is
 * left out of the new row (no such cell appears in the contract formats).
 */
function insertRowWithFormat(
  tr: Transaction,
  rect: TableRect,
  row: number
): number {
  const { map, table, tableStart } = rect;

  let rowPos = tableStart;
  for (let i = 0; i < row; i++) rowPos += table.child(i).nodeSize;

  // The reference row to inherit formatting from. Usually the row right above, and
  // only when inserting at the very top, the current first row.
  const referenceRowOffset = row > 0 ? -1 : 0;
  const referenceRow =
    table.childCount > 0 ? table.child(row + referenceRowOffset) : null;

  // `index` is derived from `col` again on every pass. Because we skip ahead by the
  // merge span, advancing the two separately would put them out of sync.
  const cells: PMNode[] = [];
  for (let col = 0; col < map.width; ) {
    const index = map.width * row + col;
    if (isCoveredFromAbove(map, row, index)) {
      col += growRowspan(tr, rect, index);
      continue;
    }
    const { node, colspan } = inheritedCell(rect, index, referenceRowOffset);
    if (node) cells.push(node);
    col += colspan;
  }

  const rowType = tableNodeTypes(table.type.schema).row;
  tr.insert(rowPos, rowType.create(inheritRowAttrs(referenceRow), cells));
  return rowPos;
}

function buildAddRowTransaction(
  state: EditorState,
  row: (rect: TableRect) => number
): Transaction | null {
  if (!pmIsInTable(state)) return null;
  const rect: TableRect = selectedRect(state);
  const tr = state.tr;
  const rowPos = insertRowWithFormat(tr, rect, row(rect));
  tr.setSelection(TextSelection.near(tr.doc.resolve(rowPos + 1)));
  return tr.scrollIntoView();
}

/**
 * Builds the transaction that adds a row right below the row holding the cursor.
 * `null` when the cursor is outside a table. The cursor moves to the first cell of the
 * new row.
 *
 * Split out of the command so it can be verified from document state alone, without an
 * editor view.
 */
export function buildAddRowAfterTransaction(
  state: EditorState
): Transaction | null {
  return buildAddRowTransaction(state, (rect) => rect.bottom);
}

/** The above version. Only the insertion position differs */
export function buildAddRowBeforeTransaction(
  state: EditorState
): Transaction | null {
  return buildAddRowTransaction(state, (rect) => rect.top);
}

/**
 * Builds the transaction that deletes the row holding the cursor (or the rows spanned
 * by a multi-row selection).
 * `null` outside a table, or when every row of the table is selected.
 *
 * The upstream `deleteRow` deletes every row the selection rectangle covers. With the
 * cursor inside a merged cell the rectangle becomes that whole cell, so all the rows it
 * covers go away with it.
 */
export function buildDeleteRowTransaction(
  state: EditorState
): Transaction | null {
  if (!pmIsInTable(state)) return null;
  const captured: Transaction[] = [];
  pmDeleteRow(state, (tr) => captured.push(tr));
  return captured[0] ?? null;
}
