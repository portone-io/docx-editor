/**
 * Newly created cells and rows inherit only the formatting of the cell or row they
 * are based on. Values that have to be recomputed from the grid (colspan, rowspan,
 * colwidth) are not inherited.
 *
 * What that comes to per attr is `docx/cloning`, which answers the same question for a
 * paragraph an edit makes out of another paragraph.
 */

import type { Node as PMNode } from "prosemirror-model";
import { CLONE_POLICIES } from "../docx/cloning";
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

/** Copies only the formatting of the reference cell. Values passed in `overrides` win */
export function inheritCellAttrs(
  cell: PMNode,
  overrides: NodeAttrs = {}
): NodeAttrs {
  return { ...CLONE_POLICIES.tableCell.attrs(cell, "copy"), ...overrides };
}

/**
 * Copies only the formatting of the reference row. `null` when there is no reference
 * row, meaning the schema defaults should be used
 */
export function inheritRowAttrs(row: PMNode | null): NodeAttrs | null {
  return row ? CLONE_POLICIES.tableRow.attrs(row, "copy") : null;
}
