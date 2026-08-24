/**
 * Grid widths (`gridCols`) are always in dxa (1/20 of a point).
 * Cell widths (`tcW`) and table widths (`tblW`) carry a value together with a unit,
 * so a width taken from the grid has to be converted into the unit that cell uses.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";
import {
  spanCount,
  type TableWidth,
  type TableWidthType,
  toTableWidth,
  widthNumber,
  withWidthNumber,
} from "../model/format";

/** The value that stands for 100% in a percentage width */
const PCT_FULL = 5000;

function isNumberArray(value: unknown): value is number[] {
  return (
    Array.isArray(value) && value.every((entry) => typeof entry === "number")
  );
}

/**
 * The grid column widths the table remembers, read from docx `w:gridCol`.
 * An empty list means the grid is unknown, so it is treated the same as an absent one
 * (no zero-width columns are created).
 */
export function tableGridCols(table: PMNode): number[] | null {
  const value: unknown = table.attrs.gridCols;
  return isNumberArray(value) && value.length > 0 ? value : null;
}

/** The sum of the grid column widths, which is the dxa width the table occupies */
export function gridTotal(gridCols: number[]): number {
  return gridCols.reduce((sum, width) => sum + width, 0);
}

/**
 * Redistributes the grid column widths proportionally so that they sum to `total`.
 *
 * Widths are integer dxa, so scaling them proportionally leaves a rounding error
 * behind. The leftover error is absorbed by the last column, making the sum exactly
 * `total`. When there is no width to distribute (the sum is 0), they are left as they
 * are.
 */
export function gridColsScaledTo(gridCols: number[], total: number): number[] {
  const current = gridTotal(gridCols);
  if (current <= 0 || total <= 0) return gridCols;

  const scaled = gridCols.map((width) => Math.round((width * total) / current));
  const drift = total - gridTotal(scaled);
  const last = scaled.length - 1;
  return scaled.map((width, at) => (at === last ? width + drift : width));
}

/**
 * Converts the width of one grid column into the unit the cell uses. A percentage is
 * a ratio of the grid total
 */
export function gridWidthInCellUnit(
  gridWidth: number,
  gridCols: number[],
  type: TableWidthType
): number {
  if (type !== "pct") return gridWidth;
  const total = gridTotal(gridCols);
  return total > 0 ? Math.round((gridWidth / total) * PCT_FULL) : gridWidth;
}

/**
 * The width of a new cell that occupies a single grid column.
 *
 * When the grid is known, one grid column is converted into the reference cell's unit;
 * when it is not, the reference cell's width is divided by its span count (using the
 * full width of a merged cell as is would make the new cell several columns wide).
 * If the reference cell records no width, none is created for the new cell either.
 */
export function cellWidthForGridCol(
  gridCols: number[] | null,
  gridCol: number,
  reference: TableWidth | null,
  referenceColspan: number
): TableWidth | null {
  if (!reference) return null;
  const gridWidth = gridCols?.[gridCol];
  if (gridCols && typeof gridWidth === "number") {
    return withWidthNumber(
      reference,
      gridWidthInCellUnit(gridWidth, gridCols, reference.type)
    );
  }
  const own = widthNumber(reference);
  return own === null
    ? reference
    : withWidthNumber(reference, Math.round(own / referenceColspan));
}

/**
 * The width of a single cell covering several grid columns.
 * It is null when the grid is unknown or the range falls outside the grid, leaving the
 * caller to pick another value.
 */
export function gridSpanWidth(
  gridCols: number[] | null,
  from: number,
  to: number,
  type: TableWidthType
): number | null {
  if (!gridCols) return null;
  const covered = gridCols.slice(from, to);
  if (covered.length === 0) return null;
  return gridWidthInCellUnit(gridTotal(covered), gridCols, type);
}

/**
 * The table width shifted by however much the grid total moved.
 * Only dxa shares its unit with the grid. A percentage is a ratio of the page width,
 * so it is left as it is.
 */
export function shiftedTableWidth(
  table: PMNode,
  gridCols: number[]
): TableWidth | null {
  const width = toTableWidth(table.attrs.tblW);
  const previous = tableGridCols(table);
  if (!width || width.type !== "dxa" || !previous) return width;
  const shifted = width.twips + gridTotal(gridCols) - gridTotal(previous);
  return { type: "dxa", twips: shifted };
}

/**
 * Writes the grid column widths back onto the table and brings the table width in line
 * with the grid total.
 *
 * `tableStart - 1` is the position of the table itself. Cell insertions and deletions
 * all happen after it, so that position never changes, but it is mapped once anyway to
 * stay safe regardless of call order.
 */
export function writeGridCols(
  tr: Transaction,
  tableStart: number,
  gridCols: number[]
): void {
  const tablePos = tr.mapping.map(tableStart - 1);
  const table = tr.doc.nodeAt(tablePos);
  if (!table) return;
  tr.setNodeMarkup(tablePos, null, {
    ...table.attrs,
    gridCols,
    tblW: shiftedTableWidth(table, gridCols),
  });
}

/**
 * Fits the cell widths to the changed grid. Every cell records the new sum of the grid
 * columns it covers.
 *
 * Only widths recorded in dxa are touched. A percentage is a ratio of the table width,
 * so it stays the same even when the grid moves, and a cell with no width (null) gets
 * none created because its original XML has to be written back as it is.
 */
export function rescaleCellWidths(
  tr: Transaction,
  tableStart: number,
  gridCols: number[]
): void {
  const tablePos = tr.mapping.map(tableStart - 1);
  const table = tr.doc.nodeAt(tablePos);
  if (!table) return;

  const map = TableMap.get(table);
  // Only widths change, so cell sizes stay the same: an earlier update never shifts
  // the position of a later cell.
  for (const pos of new Set(map.map)) {
    const cell = table.nodeAt(pos);
    const width = cell ? toTableWidth(cell.attrs.tcW) : null;
    if (!cell || !width || width.type !== "dxa") continue;

    const from = map.colCount(pos);
    const to = from + spanCount(cell.attrs.colspan);
    const spanned = gridSpanWidth(gridCols, from, to, width.type);
    if (spanned === null || spanned === width.twips) continue;
    tr.setNodeMarkup(tablePos + 1 + pos, null, {
      ...cell.attrs,
      tcW: { type: "dxa", twips: spanned },
    });
  }
}
