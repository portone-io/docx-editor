/**
 * A boundary drag moves the grid and the cell widths together exactly as a column
 * insertion does.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { bodyWidth, type PageGeometry } from "../docx/pageGeometry";
import { toTableFormat } from "../model/format";
import {
  gridTotal,
  rescaleCellWidths,
  tableGridCols,
  writeGridCols,
} from "./widths";

/**
 * The floor a column can be narrowed to (dxa).
 *
 * Word's default cell padding is 0.08 inch (115 dxa) on the left and on the right, 230
 * in total.
 * A column narrower than that cannot fit even a single character, so the cell looks
 * collapsed.
 * With a little slack added, 240 dxa (12pt, about 0.42cm) is taken as the floor.
 */
export const MIN_COLUMN_DXA = 240;

/** One point is 20 twips (dxa) */
const TWIPS_PER_PT = 20;

/** A single boundary being dragged. `col` is the index of the grid column left of it */
export interface ColumnEdge {
  /** The position where the table node sits */
  tablePos: number;
  col: number;
  /**
   * The width (dxa) the column left of the boundary wants to take. The limits are
   * re-applied here
   */
  width: number;
}

function clamp(value: number, low: number, high: number): number {
  return Math.min(Math.max(value, low), high);
}

/**
 * The floor a column can be narrowed to in this table.
 *
 * Usually `MIN_COLUMN_DXA`.
 * A table whose grid records only ratios rather than widths has a total of just a few
 * hundred dxa, which makes that floor larger than the table itself.
 * For such tables the floor is lowered to half of an evenly divided column width, so
 * that boundaries can still be dragged.
 */
function minColumnDxa(gridCols: number[]): number {
  const even = Math.floor(gridTotal(gridCols) / (2 * gridCols.length));
  return Math.min(MIN_COLUMN_DXA, Math.max(even, 1));
}

/**
 * The limit a table can be widened to (dxa).
 *
 * The body width of the paper the document names, minus the table indent (`w:tblInd`).
 * For a table already stored wider than that, its current width is the limit, so
 * dragging can only narrow it.
 */
export function maxGridTotal(table: PMNode, geometry: PageGeometry): number {
  const indentPt = toTableFormat(table.attrs.format)?.indentLeftPt ?? 0;
  const room = bodyWidth(geometry).twips - Math.round(indentPt * TWIPS_PER_PT);
  const gridCols = tableGridCols(table);
  return Math.max(room, gridCols ? gridTotal(gridCols) : 0);
}

/**
 * The grid column widths after the boundary has moved. Null when nothing changes or
 * there is no room to move into.
 *
 * At an inner boundary (`col` is not the last column) the left column receives only
 * what the right column gives up, so the grid total stays the same. At the end boundary
 * only the last column changes, so the total moves with it.
 * No column ever ends up narrower than this table's floor, and the grid total never
 * exceeds `maxTotal`.
 */
export function resizedGridCols(
  gridCols: number[],
  col: number,
  width: number,
  maxTotal: number
): number[] | null {
  const last = gridCols.length - 1;
  const current = gridCols[col];
  if (col < 0 || col > last || typeof current !== "number") return null;

  const floor = minColumnDxa(gridCols);
  const inner = col < last;
  const room = inner
    ? current + gridCols[col + 1] - floor
    : maxTotal - (gridTotal(gridCols) - current);
  if (room < floor) return null;

  const next = clamp(Math.round(width), floor, room);
  if (next === current) return null;

  return gridCols.map((entry, at) => {
    if (at === col) return next;
    if (inner && at === col + 1) return entry - (next - current);
    return entry;
  });
}

/** The table at this position. Null if what sits there is not a table */
function tableAt(doc: PMNode, pos: number): PMNode | null {
  const node = pos >= 0 && pos < doc.content.size ? doc.nodeAt(pos) : null;
  return node?.type.spec.tableRole === "table" ? node : null;
}

/**
 * Because it is one single transaction, a single undo returns everything to the state
 * before the drag.
 *
 * The geometry is the paper the open document names, which is what the table is kept
 * inside. The editor layer reads it off the state; this layer is handed it.
 */
export function buildResizeColumnTransaction(
  state: EditorState,
  edge: ColumnEdge,
  geometry: PageGeometry
): Transaction | null {
  const table = tableAt(state.doc, edge.tablePos);
  const gridCols = table ? tableGridCols(table) : null;
  if (!table || !gridCols) return null;

  const resized = resizedGridCols(
    gridCols,
    edge.col,
    edge.width,
    maxGridTotal(table, geometry)
  );
  if (!resized) return null;

  const tr = state.tr;
  writeGridCols(tr, edge.tablePos + 1, resized);
  rescaleCellWidths(tr, edge.tablePos + 1, resized);
  return tr;
}
