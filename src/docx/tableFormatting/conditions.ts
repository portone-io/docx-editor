/**
 * Which parts of a table one cell belongs to, and what its table style dresses those parts with.
 *
 * A table style says one thing for the header row, another for the first column, another for the
 * banded rows (`w:tblStylePr`), and `w:tblLook` says which of those a table takes at all. Every
 * caller that asks what a cell falls back on - the import, a table created fresh, the lines derived
 * again after the grid moved, and a cell formatting edit - asks here, so a cell is dressed the same
 * whichever of them asked.
 *
 * A conditional format writes the lines of the part it dresses rather than of every cell in it:
 * its four sides are drawn where a cell lies on the edge of that part and its inside lines between
 * the cells within it, which is the same rule the table's own lines are drawn by.
 */

import type {
  BandSizes,
  CellMargins,
  CellStyleFormat,
  CellVerticalAlign,
  InsideBorders,
  TableFormat,
  TableStyleConditions,
  TableStyleOverrideType,
} from "../../model/format";
import { ST_OnOff } from "../../ooxml/simpleTypes";
import { borderLineOfCss, wAttr } from "../../ooxml/units";
import { childByLocalName } from "../../ooxml/xml";
import {
  type CellBorderDefaults,
  type CellDefaults,
  type GridEdges,
  type GridRect,
  type GridSize,
  layerCellMargins,
  NO_BAND_SIZES,
  NO_CELL_MARGINS,
  NO_INSIDE_BORDERS,
  NO_TABLE_STYLE_CONDITIONS,
} from "./reading";

/** Which parts of its table style a table takes (`w:tblLook`, §17.4.55) */
export interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  noHBand: boolean;
  noVBand: boolean;
}

/**
 * What a table writing no `w:tblLook` at all takes: the banded rows and columns, and neither the
 * header row, the last row, nor the edge columns (§17.4.55). Switching the banding off is
 * something a table has to say, which is why both band settings are worded as `no`.
 */
export const NO_LOOK: TableLook = {
  firstRow: false,
  lastRow: false,
  firstColumn: false,
  lastColumn: false,
  noHBand: false,
  noVBand: false,
};

/**
 * The bit each part is switched on by in the hexadecimal `w:val` a Transitional document may write
 * instead of the six attributes (Part 4 §14.4.12).
 */
const LEGACY_LOOK_BITS: Readonly<Record<keyof TableLook, number>> = {
  firstRow: 0x0020,
  lastRow: 0x0040,
  firstColumn: 0x0080,
  lastColumn: 0x0100,
  noHBand: 0x0200,
  noVBand: 0x0400,
};

/** Four hexadecimal digits, which is what ST_ShortHexNumber admits */
const SHORT_HEX = /^[0-9a-fA-F]{4}$/;

/**
 * What the legacy bitmask says. A `w:val` we cannot read as one, and a table writing none at all,
 * is the bitmask `0000` (Part 4 §14.4.12), which bands the rows and columns and takes nothing else.
 */
function legacyLook(value: string | null): TableLook {
  const bits =
    value !== null && SHORT_HEX.test(value) ? Number.parseInt(value, 16) : 0;
  return {
    firstRow: (bits & LEGACY_LOOK_BITS.firstRow) !== 0,
    lastRow: (bits & LEGACY_LOOK_BITS.lastRow) !== 0,
    firstColumn: (bits & LEGACY_LOOK_BITS.firstColumn) !== 0,
    lastColumn: (bits & LEGACY_LOOK_BITS.lastColumn) !== 0,
    noHBand: (bits & LEGACY_LOOK_BITS.noHBand) !== 0,
    noVBand: (bits & LEGACY_LOOK_BITS.noVBand) !== 0,
  };
}

/**
 * Which parts of its table style this table takes.
 *
 * The six attributes are what a document writes today; where one of them is missing the legacy
 * bitmask answers for it, so a document written either way reads the same.
 */
export function readTableLook(tblPr: Element | null): TableLook {
  const el = tblPr ? childByLocalName(tblPr, "tblLook") : null;
  if (!el) return NO_LOOK;
  const legacy = legacyLook(wAttr(el, "val"));
  const written = (name: keyof TableLook): boolean =>
    ST_OnOff.parse(wAttr(el, name)) ?? legacy[name];
  return {
    firstRow: written("firstRow"),
    lastRow: written("lastRow"),
    firstColumn: written("firstColumn"),
    lastColumn: written("lastColumn"),
    noHBand: written("noHBand"),
    noVBand: written("noVBand"),
  };
}

/** One conditional format a cell takes, and the part of the grid it dresses */
export interface CellCondition {
  type: TableStyleOverrideType;
  /** Its four lines are drawn on the edge of this block, and its inside lines within it */
  region: GridRect;
}

/** Where a cell sits, and which parts of the table it therefore belongs to */
export interface CellPlacement {
  rect: GridRect;
  grid: GridSize;
  /** Lowest first, in the order §17.7.6 lays them over one another */
  conditions: readonly CellCondition[];
}

/** A band is as many rows or columns as the style said, and one where it said nothing (§17.7.6.5, §17.7.6.7) */
function bandWidth(size: number | null): number {
  return size ?? 1;
}

/** The rows a cell may be banded over: the header row and the last row are not banded with the rest */
function bandedRows(grid: GridSize, look: TableLook): [number, number] {
  return [look.firstRow ? 1 : 0, look.lastRow ? grid.rows - 1 : grid.rows];
}

function bandedColumns(grid: GridSize, look: TableLook): [number, number] {
  return [
    look.firstColumn ? 1 : 0,
    look.lastColumn ? grid.cols - 1 : grid.cols,
  ];
}

/**
 * Which band of a run of rows or columns the one starting at `at` falls in, and how far that band
 * reaches. Band 1 is the first of them, so the odd-numbered bands are band 1 and the even ones
 * band 2. null where nothing is banded there at all.
 */
function bandAt(
  at: number,
  [from, to]: [number, number],
  size: number
): { first: boolean; start: number; end: number } | null {
  if (at < from || at >= to) return null;
  const index = Math.floor((at - from) / size);
  return {
    first: index % 2 === 0,
    start: from + index * size,
    end: Math.min(from + (index + 1) * size, to),
  };
}

function rowsOf(grid: GridSize, top: number, bottom: number): GridRect {
  return { top, bottom, left: 0, right: grid.cols };
}

function columnsOf(grid: GridSize, left: number, right: number): GridRect {
  return { top: 0, bottom: grid.rows, left, right };
}

/**
 * The parts of the table this cell belongs to, lowest first.
 *
 * A cell sits in a row band by the row it starts in and in a column band by the column it starts
 * in, and it is a corner only where the table takes both the row and the column that meet there.
 */
export function cellPlacement(
  rect: GridRect,
  grid: GridSize,
  look: TableLook,
  bands: BandSizes
): CellPlacement {
  const found: CellCondition[] = [];
  const add = (type: TableStyleOverrideType, region: GridRect) => {
    found.push({ type, region });
  };
  add("wholeTable", rowsOf(grid, 0, grid.rows));

  const column = look.noVBand
    ? null
    : bandAt(rect.left, bandedColumns(grid, look), bandWidth(bands.col));
  if (column) {
    add(
      column.first ? "band1Vert" : "band2Vert",
      columnsOf(grid, column.start, column.end)
    );
  }
  const row = look.noHBand
    ? null
    : bandAt(rect.top, bandedRows(grid, look), bandWidth(bands.row));
  if (row) {
    add(
      row.first ? "band1Horz" : "band2Horz",
      rowsOf(grid, row.start, row.end)
    );
  }

  const firstRow = look.firstRow && rect.top === 0;
  const lastRow = look.lastRow && rect.bottom === grid.rows && grid.rows > 1;
  const firstCol = look.firstColumn && rect.left === 0;
  const lastCol = look.lastColumn && rect.right === grid.cols && grid.cols > 1;
  if (firstRow) add("firstRow", rowsOf(grid, 0, 1));
  if (lastRow) add("lastRow", rowsOf(grid, grid.rows - 1, grid.rows));
  if (firstCol) add("firstCol", columnsOf(grid, 0, 1));
  if (lastCol) add("lastCol", columnsOf(grid, grid.cols - 1, grid.cols));

  const corner = (
    type: TableStyleOverrideType,
    top: number,
    left: number
  ): void => {
    add(type, { top, bottom: top + 1, left, right: left + 1 });
  };
  if (firstRow && firstCol) corner("nwCell", 0, 0);
  if (firstRow && lastCol) corner("neCell", 0, grid.cols - 1);
  if (lastRow && firstCol) corner("swCell", grid.rows - 1, 0);
  if (lastRow && lastCol) corner("seCell", grid.rows - 1, grid.cols - 1);

  return { rect, grid, conditions: found };
}

/** What a table lays down for its cells to draw. What its style laid down is not in the `tblPr`, so it comes along separately */
export interface TableCellSources {
  outer: TableFormat | null;
  inside: InsideBorders;
  margins: CellMargins;
  /** The parts of its style the table takes */
  look: TableLook;
  bands: BandSizes;
  /** What the style dresses each of those parts with */
  conditions: TableStyleConditions;
  /** The style the paragraphs inside the cells resolve against. null for a table wearing none */
  styleId: string | null;
}

export const NO_CELL_SOURCES: TableCellSources = {
  outer: null,
  inside: NO_INSIDE_BORDERS,
  margins: NO_CELL_MARGINS,
  look: NO_LOOK,
  bands: NO_BAND_SIZES,
  conditions: NO_TABLE_STYLE_CONDITIONS,
  styleId: null,
};

/** Which sides of the block a cell covers lie on the edge of a part of the table, merges included */
function edgesIn(rect: GridRect, region: GridRect): GridEdges {
  return {
    top: rect.top === region.top,
    bottom: rect.bottom === region.bottom,
    left: rect.left === region.left,
    right: rect.right === region.right,
  };
}

/**
 * The lines the four sides of one cell fall back on.
 *
 * A side on the edge of the table falls on the table's outer border, and one facing another cell
 * on the table's inside line. An outer side the document wrote as `none` is passed through as it
 * is: the table switched that line off, and nothing is to bring it back.
 */
function tableBorderDefaults(
  edges: GridEdges,
  outer: TableFormat | null,
  inside: InsideBorders
): CellBorderDefaults {
  return {
    top: borderLineOfCss(
      edges.top ? (outer?.borderTop ?? null) : inside.horizontal
    ),
    bottom: borderLineOfCss(
      edges.bottom ? (outer?.borderBottom ?? null) : inside.horizontal
    ),
    left: borderLineOfCss(
      edges.left ? (outer?.borderLeft ?? null) : inside.vertical
    ),
    right: borderLineOfCss(
      edges.right ? (outer?.borderRight ?? null) : inside.vertical
    ),
  };
}

/** The lines one conditional format draws on a cell of the part it dresses, over the ones already there */
function conditionBorderDefaults(
  base: CellBorderDefaults,
  format: CellStyleFormat,
  edges: GridEdges
): CellBorderDefaults {
  const line = (onEdge: boolean, edge: string | null, within: string | null) =>
    borderLineOfCss(onEdge ? edge : within);
  return {
    top:
      line(edges.top, format.borders.top, format.inside.horizontal) ?? base.top,
    bottom:
      line(edges.bottom, format.borders.bottom, format.inside.horizontal) ??
      base.bottom,
    left:
      line(edges.left, format.borders.left, format.inside.vertical) ??
      base.left,
    right:
      line(edges.right, format.borders.right, format.inside.vertical) ??
      base.right,
  };
}

/**
 * What the cell covering this block of the grid falls back on: the table's own lines and margins,
 * with every conditional format its style dresses a part it belongs to with laid over them.
 */
export function cellDefaultsFor(
  rect: GridRect,
  grid: GridSize,
  sources: TableCellSources
): CellDefaults {
  const placement = cellPlacement(rect, grid, sources.look, sources.bands);
  const whole = { top: 0, bottom: grid.rows, left: 0, right: grid.cols };
  let borders = tableBorderDefaults(
    edgesIn(rect, whole),
    sources.outer,
    sources.inside
  );
  let margins = sources.margins;
  let background: string | null = null;
  let verticalAlign: CellVerticalAlign | null = null;
  for (const { type, region } of placement.conditions) {
    const format = sources.conditions[type];
    if (!format) continue;
    borders = conditionBorderDefaults(borders, format, edgesIn(rect, region));
    margins = layerCellMargins(margins, format.margins);
    background = format.background ?? background;
    verticalAlign = format.verticalAlign ?? verticalAlign;
  }
  return {
    borders,
    margins,
    background,
    verticalAlign,
    placement: {
      tableStyleId: sources.styleId,
      conditions: placement.conditions.map((condition) => condition.type),
    },
  };
}
