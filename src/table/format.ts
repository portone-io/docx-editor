/**
 * Newly created cells and rows inherit only the formatting of the cell or row they
 * are based on. Values that have to be recomputed from the grid (colspan, rowspan,
 * colwidth) are not inherited.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { GridRect } from "../docx/tableFormatting";

export type NodeAttrs = Record<string, unknown>;

export type { GridRect };

/** Only the parts of what `selectedRect()` returns that are actually used here */
export interface TableGridMap {
  width: number;
  height: number;
  map: number[];
  colCount(pos: number): number;
  positionAt(row: number, col: number, table: PMNode): number;
  /** The coordinates the cell at this position covers, merges included */
  findCell(pos: number): GridRect;
  /** The positions of the cells the block covers, each one listed once */
  cellsInRect(rect: GridRect): number[];
}

export interface TableRect extends GridRect {
  map: TableGridMap;
  table: PMNode;
  tableStart: number;
}

function isNumberArray(value: unknown): value is number[] {
  return (
    Array.isArray(value) && value.every((entry) => typeof entry === "number")
  );
}

/** The list of on-screen widths. `prosemirror-tables` uses it for column resizing */
export function colwidthOf(cell: PMNode): number[] | null {
  const value: unknown = cell.attrs.colwidth;
  return isNumberArray(value) ? value : null;
}

/**
 * Everything a new cell inherits.
 *
 * The content control around a cell (`sdtPrefix`, and the two locks that come out of it) is
 * deliberately left out: a new cell must not quietly come into the document carrying a copy of
 * somebody else's control, let alone that control's lock.
 */
const INHERITED_CELL_ATTRS = ["tcAttrs", "tcPr", "tcW", "format"] as const;

/** Copies only the formatting of the reference cell. Values passed in `overrides` win */
export function inheritCellAttrs(
  cell: PMNode,
  overrides: NodeAttrs = {}
): NodeAttrs {
  const attrs: NodeAttrs = {};
  for (const key of INHERITED_CELL_ATTRS) attrs[key] = cell.attrs[key];
  return { ...attrs, ...overrides };
}

/**
 * Everything a new row inherits. The row height lives inside `trPr`.
 *
 * The property exceptions (`tblPrEx`) stay behind: we never read them, so a new row simply
 * follows the table's own values.
 */
const INHERITED_ROW_ATTRS = ["trAttrs", "trPr", "format"] as const;

/**
 * Copies only the formatting of the reference row. `null` when there is no reference
 * row, meaning the schema defaults should be used
 */
export function inheritRowAttrs(row: PMNode | null): NodeAttrs | null {
  if (!row) return null;
  const attrs: NodeAttrs = {};
  for (const key of INHERITED_ROW_ATTRS) attrs[key] = row.attrs[key];
  return attrs;
}
