/**
 * The lines and the fill a cell draws because of where it sits in the grid.
 *
 * Every line of a table is drawn by its cells (see `docx/tableFormatting`), so which line a side
 * falls back on depends on where in the grid the cell sits, and so does which parts of the table
 * style dress it. A new cell inherits its neighbour's formatting, which is what a background has to
 * do but not a line: the row appended under the last one would draw the table's outer line against
 * the row above it, and deleting the last row would leave the table with no line along its bottom.
 * A row added under the header row is likewise no header row.
 *
 * So the cells of a table whose grid moved derive their display values again, along the same path
 * the import takes. What a cell wrote down itself lives in its `w:tcPr` and is read straight back
 * out of it, so only the share that came from the table can change. `cellFixes` says which cells
 * carry the wrong lines and `sameFormattingInputs` whether anything they depend on moved; the
 * deriver that runs them after an edit is `editor/plugins/tableDisplay`.
 */

import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";
import {
  type CellDefaults,
  cellDefaultsFor,
  cellMarginsOf,
  insideBordersOf,
  layerCellMargins,
  layerInsideBorders,
  NO_BAND_SIZES,
  NO_CELL_DEFAULTS,
  NO_CELL_MARGINS,
  NO_INSIDE_BORDERS,
  NO_TABLE_STYLE_CONDITIONS,
  readCellProps,
  readTableLook,
  type TableCellSources,
  tblStyleIdOf,
} from "../docx/tableFormatting";
import {
  type CellFormat,
  spanCount,
  toBandSizes,
  toCellFormat,
  toCellMargins,
  toInsideBorders,
  toTableFormat,
  toTableStyleConditions,
} from "../model/format";
import { parsePropsXml } from "../ooxml/props";
import type { NodeAttrs, TableGridMap } from "./format";

function text(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

/**
 * What this table lays down for its cells to draw.
 * The values a table style laid down are not in the `tblPr`, so the table carries them separately.
 * A node never changes once it is built, so what is read out of one is read once.
 */
export function tableCellSources(table: PMNode): TableCellSources {
  const known = sourcesByTable.get(table);
  if (known) return known;
  const tblPr = text(table.attrs.tblPr);
  const props = tblPr === null ? null : parsePropsXml(tblPr);
  const sources: TableCellSources = {
    outer: toTableFormat(table.attrs.format),
    inside: layerInsideBorders(
      toInsideBorders(table.attrs.styleInside) ?? NO_INSIDE_BORDERS,
      insideBordersOf(tblPr)
    ),
    margins: layerCellMargins(
      toCellMargins(table.attrs.styleCellMargins) ?? NO_CELL_MARGINS,
      cellMarginsOf(tblPr)
    ),
    look: readTableLook(props),
    bands: toBandSizes(table.attrs.styleBands) ?? NO_BAND_SIZES,
    conditions:
      toTableStyleConditions(table.attrs.styleConditions) ??
      NO_TABLE_STYLE_CONDITIONS,
    // The style the paragraphs inside resolve against; what it lays down for the cells is already
    // in `styleConditions`
    styleId: tblStyleIdOf(props),
  };
  sourcesByTable.set(table, sources);
  return sources;
}

const sourcesByTable = new WeakMap<PMNode, TableCellSources>();

/** What the cell at this position falls back on for everything it did not write down itself */
export function cellDefaultsAt(
  map: TableGridMap,
  pos: number,
  sources: TableCellSources
): CellDefaults {
  return cellDefaultsFor(
    map.findCell(pos),
    { rows: map.height, cols: map.width },
    sources
  );
}

const CELL_FORMAT_KEYS = [
  "borderTop",
  "borderBottom",
  "borderLeft",
  "borderRight",
  "background",
  "verticalAlign",
  "paddingTopPt",
  "paddingRightPt",
  "paddingBottomPt",
  "paddingLeftPt",
] as const;

type BorderFormatKey =
  | "borderTop"
  | "borderBottom"
  | "borderLeft"
  | "borderRight";

function sameCellFormat(a: CellFormat | null, b: CellFormat | null): boolean {
  if (!a || !b) return a === b;
  return CELL_FORMAT_KEYS.every((key) => a[key] === b[key]);
}

export interface CellFix {
  /** The position of the cell within the table's content */
  pos: number;
  attrs: NodeAttrs;
}

interface DerivedCell {
  pos: number;
  attrs: NodeAttrs;
  /** Mutable display values; they never replace the cell's preserved OOXML. */
  format: CellFormat;
  /** Only properties written directly in this cell's `w:tcPr`. */
  direct: CellFormat | null;
}

function hasDirectBorder(cell: DerivedCell, key: BorderFormatKey): boolean {
  return cell.direct?.[key] !== undefined;
}

/**
 * Suppresses the inherited side of a shared edge when its opposite is a direct cell border.
 * ECMA-376 Part 1 §17.4.66 gives the direct border precedence; leaving it as the only CSS candidate
 * also keeps one segment from repainting the whole side of an adjacent merged cell.
 */
function reconcileSharedBorder(
  first: DerivedCell,
  firstKey: BorderFormatKey,
  second: DerivedCell,
  secondKey: BorderFormatKey
): void {
  const firstIsDirect = hasDirectBorder(first, firstKey);
  const secondIsDirect = hasDirectBorder(second, secondKey);
  if (firstIsDirect === secondIsDirect) return;
  if (firstIsDirect) {
    second.format[secondKey] = "none";
  } else {
    first.format[firstKey] = "none";
  }
}

function reconcileSharedBorders(
  map: TableGridMap,
  cells: ReadonlyMap<number, DerivedCell>
): void {
  for (let row = 0; row < map.height; row += 1) {
    for (let col = 1; col < map.width; col += 1) {
      const leftPos = map.map[row * map.width + col - 1];
      const rightPos = map.map[row * map.width + col];
      if (leftPos === rightPos) continue;
      const left = cells.get(leftPos);
      const right = cells.get(rightPos);
      if (left && right) {
        reconcileSharedBorder(left, "borderRight", right, "borderLeft");
      }
    }
  }
  for (let row = 1; row < map.height; row += 1) {
    for (let col = 0; col < map.width; col += 1) {
      const topPos = map.map[(row - 1) * map.width + col];
      const bottomPos = map.map[row * map.width + col];
      if (topPos === bottomPos) continue;
      const top = cells.get(topPos);
      const bottom = cells.get(bottomPos);
      if (top && bottom) {
        reconcileSharedBorder(top, "borderBottom", bottom, "borderTop");
      }
    }
  }
}

/** The cells of this table whose display values do not match the spot they now sit in */
export function cellFixes(table: PMNode): CellFix[] {
  const map = TableMap.get(table);
  const sources = tableCellSources(table);
  const cells = new Map<number, DerivedCell>();
  // A merged cell is pointed at from every spot it covers, so each cell is looked at once
  for (const pos of new Set(map.map)) {
    const cell = table.nodeAt(pos);
    if (!cell) continue;
    const tcPr = text(cell.attrs.tcPr);
    const format = readCellProps(tcPr, cellDefaultsAt(map, pos, sources));
    cells.set(pos, {
      pos,
      attrs: cell.attrs,
      format: { ...format },
      direct: readCellProps(tcPr, NO_CELL_DEFAULTS),
    });
  }
  reconcileSharedBorders(map, cells);

  const fixes: CellFix[] = [];
  for (const cell of cells.values()) {
    const format = Object.keys(cell.format).length === 0 ? null : cell.format;
    if (sameCellFormat(format, toCellFormat(cell.attrs.format))) continue;
    fixes.push({ pos: cell.pos, attrs: { ...cell.attrs, format } });
  }
  return fixes;
}

function sameSpans(a: PMNode, b: PMNode): boolean {
  return (
    spanCount(a.attrs.colspan) === spanCount(b.attrs.colspan) &&
    spanCount(a.attrs.rowspan) === spanCount(b.attrs.rowspan)
  );
}

function sameRow(a: PMNode, b: PMNode): boolean {
  return (
    a.childCount === b.childCount &&
    a.children.every((cell, index) => sameSpans(cell, b.child(index)))
  );
}

/** The grid is all a cell's lines depend on, so a table whose text alone was edited is left alone */
function sameGrid(a: PMNode, b: PMNode): boolean {
  return (
    a.childCount === b.childCount &&
    a.children.every((row, index) => sameRow(row, b.child(index)))
  );
}

function sameCellFormattingInputs(a: PMNode, b: PMNode): boolean {
  return a.children.every((row, rowIndex) =>
    row.children.every(
      (cell, cellIndex) =>
        cell.attrs.tcPr === b.child(rowIndex).child(cellIndex).attrs.tcPr
    )
  );
}

/** Whether everything the cells' lines are derived from reads the same in both tables */
export function sameFormattingInputs(a: PMNode, b: PMNode): boolean {
  return (
    sameGrid(a, b) &&
    a.attrs.tblPr === b.attrs.tblPr &&
    a.attrs.format === b.attrs.format &&
    a.attrs.styleInside === b.attrs.styleInside &&
    a.attrs.styleCellMargins === b.attrs.styleCellMargins &&
    a.attrs.styleConditions === b.attrs.styleConditions &&
    a.attrs.styleBands === b.attrs.styleBands &&
    sameCellFormattingInputs(a, b)
  );
}
