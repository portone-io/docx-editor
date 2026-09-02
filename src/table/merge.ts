/**
 * The grid math and moving the content around are left to `prosemirror-tables`; only
 * the formatting of the resulting cells is rewritten.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import {
  mergeCells as pmMergeCells,
  splitCell as pmSplitCell,
  selectedRect,
  TableMap,
} from "prosemirror-tables";
import {
  spanCount,
  type TableWidth,
  type TableWidthType,
  toTableWidth,
  widthNumber,
  withWidthNumber,
} from "../model/format";
import { transactionAllowed } from "../schema/locks";
import { inheritCellAttrs, type TableRect } from "./format";
import { cellWidthForGridCol, gridSpanWidth, tableGridCols } from "./widths";

interface TableAfterChange {
  tablePos: number;
  table: PMNode;
  map: TableMap;
}

/**
 * Finds the same table again in the document once the operation is done. Its position
 * and its grid may have changed in the meantime
 */
function tableAfter(
  tr: Transaction,
  tableStart: number
): TableAfterChange | null {
  const tablePos = tr.mapping.map(tableStart - 1);
  const table = tr.doc.nodeAt(tablePos);
  if (!table) return null;
  return { tablePos, table, map: TableMap.get(table) };
}

function cellAt(rect: TableRect, row: number, col: number): PMNode | null {
  const pos = rect.map.map[row * rect.map.width + col];
  return pos == null ? null : rect.table.nodeAt(pos);
}

/**
 * Whether several cells can be merged into one.
 * The selection has to be a rectangle spanning more than one cell, no merged cell may stick out
 * past that rectangle, and the lock guard has to let the merge through: a locked cell may not be
 * swallowed by another (`schema/locks`).
 *
 * The two queries are defined from the very transaction the commands run, so they cannot drift
 * from them. They live here rather than beside the commands because `./commands` imports these
 * builders, and the other way round would turn that import around.
 */
export function canMergeCells(state: EditorState): boolean {
  const tr = buildMergeCellsTransaction(state);
  return tr !== null && transactionAllowed(tr, state);
}

/**
 * Whether a split is possible. The cursor has to sit inside a cell that is merged horizontally or
 * vertically, and the lock guard has to let the split through: the cells a split makes would each
 * carry the original's lock, which is a lock planted in places it was never put (`schema/locks`).
 */
export function canSplitCell(state: EditorState): boolean {
  const tr = buildSplitCellTransaction(state);
  return tr !== null && transactionAllowed(tr, state);
}

/**
 * The sum of the widths recorded by the cells the merge range covers.
 * If a cell has no width or one with a different unit is mixed in, the sum cannot be
 * trusted, so the result is null.
 */
function coveredWidthSum(rect: TableRect, type: TableWidthType): number | null {
  let sum = 0;
  let previous: number | null = null;
  for (let col = rect.left; col < rect.right; col++) {
    const pos = rect.map.map[rect.top * rect.map.width + col];
    // A horizontally merged cell points at the same position from every column it
    // covers, so count it only once
    if (pos === previous) continue;
    previous = pos;
    const cell = pos == null ? null : rect.table.nodeAt(pos);
    const width = cell ? toTableWidth(cell.attrs.tcW) : null;
    const value = width ? widthNumber(width) : null;
    if (!width || width.type !== type || value === null) return null;
    sum += value;
  }
  return sum;
}

/**
 * The width of a merged cell that ends up covering several columns.
 *
 * Leaving the top-left cell's width as it is makes `w:tcW` smaller than the grid the
 * cell actually covers, which throws the table out of alignment in Word. So when the
 * grid is known, the sum of the covered column widths is used, and when it is not, the
 * sum of the widths recorded by the covered cells, expressed in the top-left cell's
 * unit.
 * If neither can be worked out, the top-left cell's width is left as it is.
 */
function mergedCellWidth(rect: TableRect, source: PMNode): TableWidth | null {
  const reference = toTableWidth(source.attrs.tcW);
  if (!reference) return null;
  const spanned =
    gridSpanWidth(
      tableGridCols(rect.table),
      rect.left,
      rect.right,
      reference.type
    ) ?? coveredWidthSum(rect, reference.type);
  return spanned === null ? reference : withWidthNumber(reference, spanned);
}

function writeMergedFormat(
  tr: Transaction,
  rect: TableRect,
  source: PMNode
): void {
  const after = tableAfter(tr, rect.tableStart);
  if (!after) return;
  const { tablePos, table, map } = after;
  const pos = map.map[rect.top * map.width + rect.left];
  const merged = table.nodeAt(pos);
  if (!merged) return;
  tr.setNodeMarkup(tablePos + 1 + pos, null, {
    ...merged.attrs,
    ...inheritCellAttrs(source, { tcW: mergedCellWidth(rect, source) }),
  });
}

/**
 * `null` if the grid does not allow the merge. The lock is not asked here: what the guard
 * would turn down is what `canMergeCells` and the command are for, and the transaction is
 * what they ask about.
 *
 * The size of the merged cell (colspan, rowspan, colwidth) keeps whatever upstream
 * decided, and its formatting comes from the top-left cell.
 */
export function buildMergeCellsTransaction(
  state: EditorState
): Transaction | null {
  if (!pmMergeCells(state)) return null;
  const rect: TableRect = selectedRect(state);
  const source = cellAt(rect, rect.top, rect.left);

  const captured: Transaction[] = [];
  pmMergeCells(state, (tr) => captured.push(tr));
  const tr = captured[0];
  if (!tr || !source) return null;

  writeMergedFormat(tr, rect, source);
  return tr;
}

/**
 * Writes the width of its own grid column onto every cell produced by the split.
 *
 * Upstream builds the new cells out of the original cell's attrs, so the content control the
 * original sat inside would be copied onto each of them, its lock along with it. Only the cell
 * still standing in the original's own spot keeps either.
 */
function writeSplitFormats(
  tr: Transaction,
  rect: TableRect,
  source: PMNode
): void {
  const after = tableAfter(tr, rect.tableStart);
  if (!after) return;
  const { tablePos, table, map } = after;
  const gridCols = tableGridCols(table);
  const width = toTableWidth(source.attrs.tcW);
  const colspan = spanCount(source.attrs.colspan);

  for (let row = rect.top; row < rect.bottom; row++) {
    for (let col = rect.left; col < rect.right; col++) {
      const pos = map.map[row * map.width + col];
      const cell = pos == null ? null : table.nodeAt(pos);
      if (!cell) continue;
      // Only formatting changes, so cell sizes stay the same: an earlier update never
      // shifts the position of a later cell.
      const isOriginalSpot = row === rect.top && col === rect.left;
      tr.setNodeMarkup(tablePos + 1 + pos, null, {
        ...cell.attrs,
        ...inheritCellAttrs(source, {
          tcW: cellWidthForGridCol(gridCols, col, width, colspan),
          sdtPrefix: isOriginalSpot ? source.attrs.sdtPrefix : null,
          sdtContentsLocked: isOriginalSpot
            ? source.attrs.sdtContentsLocked
            : false,
          sdtDeletionLocked: isOriginalSpot
            ? source.attrs.sdtDeletionLocked
            : false,
        }),
      });
    }
  }
}

/**
 * `null` for a cell that is not merged, or outside a table. As with the merge, the lock is asked
 * of the finished transaction by `canSplitCell` and the command rather than here.
 */
export function buildSplitCellTransaction(
  state: EditorState
): Transaction | null {
  if (!pmSplitCell(state)) return null;
  const rect: TableRect = selectedRect(state);
  const source = cellAt(rect, rect.top, rect.left);

  const captured: Transaction[] = [];
  pmSplitCell(state, (tr) => captured.push(tr));
  const tr = captured[0];
  if (!tr || !source) return null;

  writeSplitFormats(tr, rect, source);
  return tr;
}
