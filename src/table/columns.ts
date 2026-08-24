/**
 * `prosemirror-tables`' `addColumn` gets the grid math right but creates the new cells
 * without any formatting. So only its skeleton is followed here, with inheritance of
 * the reference cell's formatting and an update of the grid column widths added on top.
 * Deletion is left to `deleteColumn`, and only the deleted columns are dropped from the
 * grid.
 *
 * Adding a column leaves the grid total the same as it was before: widening the table
 * by the new column's width would push it past the page.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  type EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import {
  addColSpan,
  deleteColumn as pmDeleteColumn,
  isInTable as pmIsInTable,
  selectedRect,
  TableMap,
  tableNodeTypes,
} from "prosemirror-tables";
import { spanCount, toTableWidth } from "../model/format";
import {
  colwidthOf,
  inheritCellAttrs,
  type NodeAttrs,
  type TableGridMap,
  type TableRect,
} from "./format";
import {
  cellWidthForGridCol,
  gridColsScaledTo,
  gridTotal,
  rescaleCellWidths,
  tableGridCols,
  writeGridCols,
} from "./widths";

/**
 * The shape `addColSpan` requires. Restated here because `prosemirror-tables` does not
 * export the type
 */
interface SpanAttrs extends NodeAttrs {
  colspan: number;
  rowspan: number;
  colwidth: number[] | null;
}

function spanAttrsOf(cell: PMNode): SpanAttrs {
  return {
    ...cell.attrs,
    colspan: spanCount(cell.attrs.colspan),
    rowspan: spanCount(cell.attrs.rowspan),
    colwidth: colwidthOf(cell),
  };
}

/**
 * Slices out just the one entry of the reference cell's `colwidth` that belongs to the
 * new column
 */
function columnSlice(cell: PMNode, offset: number): number[] | null {
  const value = colwidthOf(cell);
  if (!value) return null;
  const entry = value[offset] ?? value[0];
  return typeof entry === "number" ? [entry] : null;
}

function newColumnCellAttrs(
  rect: TableRect,
  referenceCell: PMNode,
  referencePos: number,
  referenceGridCol: number
): NodeAttrs {
  return inheritCellAttrs(referenceCell, {
    colspan: 1,
    rowspan: 1,
    colwidth: columnSlice(
      referenceCell,
      referenceGridCol - rect.map.colCount(referencePos)
    ),
    tcW: cellWidthForGridCol(
      tableGridCols(rect.table),
      referenceGridCol,
      toTableWidth(referenceCell.attrs.tcW),
      spanCount(referenceCell.attrs.colspan)
    ),
  });
}

/**
 * If this position points at the same cell as the one to its left, it is a position
 * crossed by a horizontally merged cell
 */
function isCrossedFromLeft(
  map: TableGridMap,
  col: number,
  index: number
): boolean {
  return col > 0 && col < map.width && map.map[index - 1] === map.map[index];
}

/**
 * Grows by one the width of a merged cell that crosses the new column.
 * Returns how many rows that cell covers (that many have to be skipped).
 */
function growColspan(
  tr: Transaction,
  rect: TableRect,
  col: number,
  index: number
): number {
  const { map, table, tableStart } = rect;
  const pos = map.map[index];
  const cell = table.nodeAt(pos);
  if (!cell) return 1;

  const attrs = spanAttrsOf(cell);
  tr.setNodeMarkup(
    tr.mapping.map(tableStart + pos),
    null,
    addColSpan(attrs, col - map.colCount(pos))
  );
  return attrs.rowspan;
}

/**
 * When the table width is a share of the page, the table stays inside the page even as
 * the grid grows
 */
function widthIsPageShare(table: PMNode): boolean {
  return toTableWidth(table.attrs.tblW)?.type === "pct";
}

/**
 * Inserts the reference column's width as one new entry at the new column's slot in the
 * grid.
 *
 * Unless the table width is a share of the page, the grid total is the table's width.
 * So after the insertion the whole grid is scaled down proportionally to bring the total
 * back to what it was before, and the cell widths are fitted to the new grid as well.
 */
function insertGridCol(
  tr: Transaction,
  rect: TableRect,
  col: number,
  referenceGridCol: number
): void {
  const gridCols = tableGridCols(rect.table);
  if (!gridCols) return;

  const at = Math.min(col, gridCols.length);
  const inserted =
    gridCols[referenceGridCol] ?? gridCols[gridCols.length - 1] ?? 0;
  const grown = [...gridCols.slice(0, at), inserted, ...gridCols.slice(at)];

  if (widthIsPageShare(rect.table)) {
    writeGridCols(tr, rect.tableStart, grown);
    return;
  }

  const shared = gridColsScaledTo(grown, gridTotal(gridCols));
  writeGridCols(tr, rect.tableStart, shared);
  rescaleCellWidths(tr, rect.tableStart, shared);
}

/**
 * Inserts a column at grid column `col` that inherits the formatting of the column to
 * its left.
 * If `col` equals the table's column count, the column is appended at the right edge.
 *
 * Unlike a row insertion, one cell goes into every row, so from the second row on the
 * positions have already been pushed along by the earlier insertions. That is why the
 * position is mapped again for each insertion.
 */
function insertColumnWithFormat(
  tr: Transaction,
  rect: TableRect,
  col: number
): void {
  const { map, table, tableStart } = rect;
  // The reference column to inherit formatting from. Usually the column immediately to
  // the left, and only when inserting at the very front, the current first column.
  const referenceOffset = col > 0 ? -1 : 0;
  const referenceGridCol = col + referenceOffset;

  for (let row = 0; row < map.height; row++) {
    const index = row * map.width + col;
    if (isCrossedFromLeft(map, col, index)) {
      row += growColspan(tr, rect, col, index) - 1;
      continue;
    }

    const referencePos = map.map[index + referenceOffset];
    const referenceCell =
      referencePos != null ? table.nodeAt(referencePos) : null;
    const type = referenceCell
      ? referenceCell.type
      : tableNodeTypes(table.type.schema).cell;
    const attrs = referenceCell
      ? newColumnCellAttrs(rect, referenceCell, referencePos, referenceGridCol)
      : null;
    const node = type.createAndFill(attrs);
    if (!node) continue;
    tr.insert(
      tr.mapping.map(tableStart + map.positionAt(row, col, table)),
      node
    );
  }

  insertGridCol(tr, rect, col, referenceGridCol);
}

function buildAddColumnTransaction(
  state: EditorState,
  column: (rect: TableRect) => number
): Transaction | null {
  if (!pmIsInTable(state)) return null;
  const rect: TableRect = selectedRect(state);
  const col = column(rect);
  const tr = state.tr;
  insertColumnWithFormat(tr, rect, col);

  // The cursor position is looked up again in the document once the insertions are
  // done. Mapping a pre-insertion coordinate would point at the original cell, now
  // pushed to the right, rather than at the new one.
  const tablePos = tr.mapping.map(rect.tableStart - 1);
  const table = tr.doc.nodeAt(tablePos);
  if (table) {
    const map = TableMap.get(table);
    const row = Math.min(rect.top, map.height - 1);
    const cellPos = map.map[row * map.width + col];
    tr.setSelection(TextSelection.near(tr.doc.resolve(tablePos + 1 + cellPos)));
  }
  return tr.scrollIntoView();
}

/**
 * Builds the transaction that adds a column right of the column holding the cursor.
 * `null` when the cursor is outside a table. The cursor moves into a cell of the new
 * column.
 */
export function buildAddColumnAfterTransaction(
  state: EditorState
): Transaction | null {
  return buildAddColumnTransaction(state, (rect) => rect.right);
}

/** The left-hand version. Only the insertion position differs */
export function buildAddColumnBeforeTransaction(
  state: EditorState
): Transaction | null {
  return buildAddColumnTransaction(state, (rect) => rect.left);
}

/**
 * Builds the transaction that deletes the column holding the cursor (or the columns
 * spanned by a multi-column selection).
 * `null` outside a table, or when every column of the table is selected.
 */
export function buildDeleteColumnTransaction(
  state: EditorState
): Transaction | null {
  if (!pmIsInTable(state)) return null;
  const rect: TableRect = selectedRect(state);
  // `deleteColumn` overwrites the rect as it loops, so the values needed are read up
  // front.
  const { left, right, tableStart } = rect;
  const gridCols = tableGridCols(rect.table);

  const captured: Transaction[] = [];
  pmDeleteColumn(state, (tr) => captured.push(tr));
  const tr = captured[0];
  if (!tr) return null;

  if (gridCols) {
    writeGridCols(tr, tableStart, [
      ...gridCols.slice(0, left),
      ...gridCols.slice(right),
    ]);
  }
  return tr;
}
