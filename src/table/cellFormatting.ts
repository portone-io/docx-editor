/**
 * A job applies to the cell holding the caret, or to every cell a cell selection covers.
 * All XML handling belongs to `docx/tableFormatting`; this file only decides which cells take part
 * and which of their sides a border preset writes.
 *
 * A cell already in the state the job wants is left untouched, so its original XML survives.
 * That is also what tells the menu whether an entry has anything to do: following the ProseMirror
 * convention, calling a command without `dispatch` only reports whether it would change something.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { isInTable as pmIsInTable, selectedRect } from "prosemirror-tables";
import {
  ALL_CELL_SIDES,
  type CellFormatEdit,
  type CellSide,
  editCellProps,
} from "../docx/tableFormatting";
import {
  type CellFormat,
  type CellVerticalAlign,
  NO_FILL,
  toCellFormat,
} from "../model/format";
import { isLockedCell, transactionAllowed } from "../schema/locks";
import type { TableCommand } from "./commands";
import type { NodeAttrs, TableRect } from "./format";
import { cellDefaultsAt, tableCellSources } from "./gridBorders";

/**
 * Which sides of the selected block a preset writes.
 * `all` and `none` speak for every side of every selected cell, while `outer` writes only the sides
 * that lie on the edge of the block, leaving the lines inside it as they are (the same as Word).
 */
export type CellBorderPreset = "all" | "outer" | "none";

function text(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

/** The sides of this cell that lie on the edge of the selected block */
function outerSides(rect: TableRect, pos: number): CellSide[] {
  const cell = rect.map.findCell(pos);
  const sides: CellSide[] = [];
  if (cell.top === rect.top) sides.push("top");
  if (cell.bottom === rect.bottom) sides.push("bottom");
  if (cell.left === rect.left) sides.push("left");
  if (cell.right === rect.right) sides.push("right");
  return sides;
}

function borderEdit(
  preset: CellBorderPreset,
  rect: TableRect,
  pos: number
): CellFormatEdit {
  return {
    kind: "borders",
    line: preset === "none" ? "none" : "single",
    sides: preset === "outer" ? outerSides(rect, pos) : ALL_CELL_SIDES,
  };
}

interface CellChange {
  pos: number;
  attrs: NodeAttrs;
}

/**
 * The new attributes of every selected cell the job changes something in.
 * Only attributes change, so no cell ever moves and the positions stay good for the whole
 * transaction.
 */
function planChanges(
  rect: TableRect,
  edit: (rect: TableRect, pos: number) => CellFormatEdit
): CellChange[] {
  // The lines and the padding a cell gets from its table are derived afresh, from where the cell
  // sits in the grid, so they survive an edit to the cell's own formatting
  const sources = tableCellSources(rect.table);
  const changes: CellChange[] = [];
  for (const pos of rect.map.cellsInRect(rect)) {
    const cell = rect.table.nodeAt(pos);
    if (!cell) continue;
    const next = editCellProps(
      text(cell.attrs.tcPr),
      edit(rect, pos),
      cellDefaultsAt(rect.map, pos, sources),
      sources.margins
    );
    if (!next) continue;
    changes.push({
      pos,
      attrs: { ...cell.attrs, tcPr: next.tcPr, format: next.format },
    });
  }
  return changes;
}

function cellFormatCommand(
  edit: (rect: TableRect, pos: number) => CellFormatEdit
): TableCommand {
  return (state, dispatch) => {
    if (!pmIsInTable(state)) return false;
    const rect: TableRect = selectedRect(state);
    if (
      rect.map
        .cellsInRect(rect)
        .some((pos) => isLockedCell(rect.table.nodeAt(pos)))
    ) {
      return false;
    }
    const changes = planChanges(rect, edit);
    if (changes.length === 0) return false;
    const transaction = applyChanges(state.tr, rect.tableStart, changes);
    if (!transactionAllowed(transaction, state.doc)) return false;
    dispatch?.(transaction);
    return true;
  };
}

export type MixedCellValue<T> = T | "mixed" | null;

export interface ActiveCellPadding {
  top: MixedCellValue<number>;
  right: MixedCellValue<number>;
  bottom: MixedCellValue<number>;
  left: MixedCellValue<number>;
}

function shared<T>(values: readonly (T | null)[]): MixedCellValue<T> {
  const [first, ...rest] = values;
  if (first === undefined) return null;
  return rest.every((value) => value === first) ? first : "mixed";
}

/** The effective vertical alignment of the selected cells. An omitted OOXML value means top. */
export function activeCellVerticalAlign(
  state: EditorState
): MixedCellValue<CellVerticalAlign> {
  return shared(
    cellFormats(state).map((format) => format?.verticalAlign ?? "top")
  );
}

/** The effective padding on each side of the selected cells, with mixed sides reported separately. */
export function activeCellPadding(state: EditorState): ActiveCellPadding {
  const formats = cellFormats(state);
  return {
    top: shared(formats.map((format) => format?.paddingTopPt ?? null)),
    right: shared(formats.map((format) => format?.paddingRightPt ?? null)),
    bottom: shared(formats.map((format) => format?.paddingBottomPt ?? null)),
    left: shared(formats.map((format) => format?.paddingLeftPt ?? null)),
  };
}

/** Places content at the top, center, or bottom of every selected cell. */
export function setCellVerticalAlign(align: CellVerticalAlign): TableCommand {
  return cellFormatCommand(() => ({ kind: "verticalAlign", align }));
}

/** Writes direct padding only for the sides supplied, leaving every omitted side unchanged. */
export function setCellPadding(
  values: Partial<Record<CellSide, number>>
): TableCommand {
  return cellFormatCommand(() => ({ kind: "padding", values }));
}

function applyChanges(
  tr: Transaction,
  tableStart: number,
  changes: readonly CellChange[]
): Transaction {
  for (const change of changes) {
    tr.setNodeMarkup(tableStart + change.pos, null, change.attrs);
  }
  return tr;
}

/** Draws (or clears) the lines of the selected cells. */
export function setCellBorders(preset: CellBorderPreset): TableCommand {
  return cellFormatCommand((rect, pos) => borderEdit(preset, rect, pos));
}

/**
 * Colors the visible lines of the selected cells as `#RRGGBB`. Inherited table lines become direct
 * cell overrides only on the sides that are visible. Null resets the resulting color to `auto`.
 */
export function setCellBorderColor(hex: string | null): TableCommand {
  return cellFormatCommand(() => ({ kind: "borderColor", hex }));
}

/** Fills the selected cells with `#RRGGBB`. Null takes the fill away */
export function setCellBackground(hex: string | null): TableCommand {
  return cellFormatCommand(() => ({ kind: "background", hex }));
}

/** The cells the current selection acts on. Empty outside a table */
function selectedCells(state: EditorState): PMNode[] {
  if (!pmIsInTable(state)) return [];
  const rect: TableRect = selectedRect(state);
  const cells: PMNode[] = [];
  for (const pos of rect.map.cellsInRect(rect)) {
    const cell = rect.table.nodeAt(pos);
    if (cell) cells.push(cell);
  }
  return cells;
}

function cellFormats(state: EditorState): (CellFormat | null)[] {
  return selectedCells(state).map((cell) => toCellFormat(cell.attrs.format));
}

/** Whether direct cell formatting can act on the whole current cell selection. */
export function canSetCellFormatting(state: EditorState): boolean {
  const cells = selectedCells(state);
  return cells.length > 0 && cells.every((cell) => !isLockedCell(cell));
}

/** Null when the selected cells do not share one value. The palette shows that as "nothing picked" */
function sharedValue(values: readonly (string | null)[]): string | null {
  const [first, ...rest] = values;
  if (first === undefined) return null;
  return rest.every((value) => value === first) ? first : null;
}

/** The color of a border display value such as `0.5pt solid #999999`. Null for a side that draws nothing */
function borderColorOf(border: string | undefined): string | null {
  if (border === undefined || border === "none") return null;
  const color = border.slice(border.lastIndexOf(" ") + 1);
  return color.startsWith("#") ? color : null;
}

/** The fill shared by the selected cells. Null when they differ or none is filled */
export function activeCellBackground(state: EditorState): string | null {
  const shared = sharedValue(
    cellFormats(state).map((format) => format?.background ?? null)
  );
  // A cell whose document says to paint nothing has no fill for the palette to point at
  return shared === NO_FILL ? null : shared;
}

/**
 * The color shared by every line the selected cells draw.
 * The lines are looked at as they appear on screen, so a color the table's inside lines gave is
 * counted too. Null when the lines differ, or when none is drawn.
 */
export function activeCellBorderColor(state: EditorState): string | null {
  const colors = cellFormats(state).flatMap((format) =>
    [
      format?.borderTop,
      format?.borderBottom,
      format?.borderLeft,
      format?.borderRight,
    ]
      .map(borderColorOf)
      .filter((color): color is string => color !== null)
  );
  return sharedValue(colors);
}

/**
 * Whether the border color has anything visible to act on in the current cell selection.
 */
export function canSetCellBorderColor(state: EditorState): boolean {
  return (
    canSetCellFormatting(state) &&
    cellFormats(state).some((format) =>
      [
        format?.borderTop,
        format?.borderBottom,
        format?.borderLeft,
        format?.borderRight,
      ].some((border) => border !== undefined && border !== "none")
    )
  );
}
