/**
 * Reads table formatting for display. Table borders and margins are folded into cell values so CSS
 * collapsed-border resolution matches OOXML inheritance without changing exported XML.
 */

import type {
  CellFormat,
  CellMargins,
  CellVerticalAlign,
  InsideBorders,
  RowFormat,
  RowHeight,
  TableFormat,
  TableWidth,
} from "../../model/format";
import {
  ALIGN_BY_JC,
  borderSide,
  childValue,
  isOn,
  shadingOf,
  toNumber,
  twipsToPt,
  wAttr,
} from "../../ooxml/units";
import { childByLocalName } from "../../ooxml/xml";
import { parsePropsXml } from "../propsXml";

export type { CellMargins, InsideBorders } from "../../model/format";

/**
 * The width written down by `<w:tblW>` or `<w:tcW>`.
 * If we cannot make out what it means it is null, and such a width goes back out unchanged on export.
 *
 * `auto` and `nil` leave the width to the layout, so whatever number stands beside them says nothing.
 * A percentage may be written in fiftieths of a percent, as in `w:w="2500"`, or as `w:w="50%"`.
 * Both have to be gathered into the same unit, or the table collapses into a thin strip on screen.
 */
export function readTableWidth(
  parent: Element | null,
  name: "tblW" | "tcW"
): TableWidth | null {
  const el = parent ? childByLocalName(parent, name) : null;
  if (!el) return null;
  const type = wAttr(el, "type") ?? "dxa";
  if (type === "auto") return { type: "auto" };
  if (type === "nil") return { type: "nil" };

  const raw = wAttr(el, "w");
  const value = toNumber(raw);
  if (value === null) return null;
  if (type === "dxa") return { type: "dxa", twips: value };
  if (type === "pct") {
    const isPercentText = raw?.endsWith("%") === true;
    return {
      type: "pct",
      fiftieths: isPercentText ? Math.round(value * 50) : value,
    };
  }
  // A unit we do not know at all
  return null;
}

/** The sequence of column widths (dxa) set by `<w:tblGrid>`. Revision markup (tblGridChange) is skipped */
export function readGridCols(tblGrid: Element | null): number[] {
  if (!tblGrid) return [];
  const cols: number[] = [];
  for (const child of Array.from(tblGrid.children)) {
    if (child.localName !== "gridCol") continue;
    const w = toNumber(wAttr(child, "w"));
    if (w !== null) cols.push(w);
  }
  return cols;
}

export const NO_INSIDE_BORDERS: InsideBorders = {
  horizontal: null,
  vertical: null,
};

export function readInsideBorders(tblPr: Element | null): InsideBorders {
  const borders = tblPr ? childByLocalName(tblPr, "tblBorders") : null;
  return {
    horizontal: borderSide(borders, "insideH"),
    vertical: borderSide(borders, "insideV"),
  };
}

/**
 * Lays the inside lines a table wrote down on top of the ones its style laid down.
 * A side the table wrote as `none` wins as well: a document that switched a line off is not to have
 * the style's line come back.
 */
export function layerInsideBorders(
  style: InsideBorders,
  direct: InsideBorders
): InsideBorders {
  return {
    horizontal: direct.horizontal ?? style.horizontal,
    vertical: direct.vertical ?? style.vertical,
  };
}

export const NO_CELL_MARGINS: CellMargins = {
  topPt: null,
  rightPt: null,
  bottomPt: null,
  leftPt: null,
};

/**
 * One side of a margins element. Its width is written as `w:w`, in the unit `w:type` names.
 * `nil` is a margin of nothing; a unit we cannot turn into points reads as nothing written down, so
 * the side keeps the margin the level above laid down.
 */
function marginSide(margins: Element, names: readonly string[]): number | null {
  for (const name of names) {
    const el = childByLocalName(margins, name);
    if (!el) continue;
    const type = wAttr(el, "type") ?? "dxa";
    if (type === "nil") return 0;
    return type === "dxa" ? twipsToPt(wAttr(el, "w")) : null;
  }
  return null;
}

/** The margins `w:tblCellMar` (in a tblPr) or `w:tcMar` (in a tcPr) writes down, one side at a time */
export function readCellMarginsOf(
  parent: Element | null,
  name: "tblCellMar" | "tcMar"
): CellMargins {
  const el = parent ? childByLocalName(parent, name) : null;
  if (!el) return NO_CELL_MARGINS;
  return {
    topPt: marginSide(el, ["top"]),
    rightPt: marginSide(el, ["end", "right"]),
    bottomPt: marginSide(el, ["bottom"]),
    leftPt: marginSide(el, ["start", "left"]),
  };
}

/**
 * Lays the margins a table wrote down on top of the ones its style laid down.
 * A side the table says nothing about keeps the style's margin, which is where most documents get
 * their cell padding from.
 */
export function layerCellMargins(
  style: CellMargins,
  direct: CellMargins
): CellMargins {
  return {
    topPt: direct.topPt ?? style.topPt,
    rightPt: direct.rightPt ?? style.rightPt,
    bottomPt: direct.bottomPt ?? style.bottomPt,
    leftPt: direct.leftPt ?? style.leftPt,
  };
}

/** The lines a cell falls back on, one per side, already resolved for its spot in the grid */
export interface CellBorderDefaults {
  top: string | null;
  bottom: string | null;
  left: string | null;
  right: string | null;
}

export const NO_BORDER_DEFAULTS: CellBorderDefaults = {
  top: null,
  bottom: null,
  left: null,
  right: null,
};

/** Which sides of a cell lie on the edge of the table's grid */
export interface GridEdges {
  top: boolean;
  bottom: boolean;
  left: boolean;
  right: boolean;
}

/** A block of grid coordinates. `bottom` and `right` point one past the last cell, not at it */
export interface GridRect {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export interface GridSize {
  rows: number;
  cols: number;
}

/**
 * Which sides of the block a cell covers lie on the edge of the grid, merges included.
 * Everything that works out a cell's lines asks this first, whether the cell is being imported,
 * created fresh, or looked up again in a table on screen.
 */
export function gridEdgesOf(rect: GridRect, grid: GridSize): GridEdges {
  return {
    top: rect.top === 0,
    bottom: rect.bottom === grid.rows,
    left: rect.left === 0,
    right: rect.right === grid.cols,
  };
}

/**
 * The lines the four sides of one cell fall back on.
 *
 * A side on the edge of the grid falls on the table's outer border, and one facing another cell on
 * the table's inside line. An outer side the document wrote as `none` is passed through as it is:
 * the table switched that line off, and nothing is to bring it back.
 */
export function cellBorderDefaults(
  edges: GridEdges,
  outer: TableFormat | null,
  inside: InsideBorders
): CellBorderDefaults {
  return {
    top: edges.top ? (outer?.borderTop ?? null) : inside.horizontal,
    bottom: edges.bottom ? (outer?.borderBottom ?? null) : inside.horizontal,
    left: edges.left ? (outer?.borderLeft ?? null) : inside.vertical,
    right: edges.right ? (outer?.borderRight ?? null) : inside.vertical,
  };
}

/** The name of the table style `w:tblStyle` points at. null if it points at none */
export function tblStyleIdOf(tblPr: Element | null): string | null {
  return tblPr ? childValue(tblPr, "tblStyle") : null;
}

export function readTableFormat(tblPr: Element | null): TableFormat | null {
  if (!tblPr) return null;
  const format: TableFormat = {};
  const borders = childByLocalName(tblPr, "tblBorders");
  const top = borderSide(borders, "top");
  if (top) format.borderTop = top;
  const bottom = borderSide(borders, "bottom");
  if (bottom) format.borderBottom = bottom;
  const left = borderSide(borders, "left") ?? borderSide(borders, "start");
  if (left) format.borderLeft = left;
  const right = borderSide(borders, "right") ?? borderSide(borders, "end");
  if (right) format.borderRight = right;

  const background = shadingOf(tblPr);
  if (background) format.background = background;

  const align = ALIGN_BY_JC[childValue(tblPr, "jc") ?? ""];
  if (align) format.align = align;

  // An indent of 0 states that the table stands at the margin, which is not the same as the table
  // saying nothing about where it stands. A negative one pushes it left of the margin
  const tblInd = childByLocalName(tblPr, "tblInd");
  const indentLeftPt = tblInd ? twipsToPt(wAttr(tblInd, "w")) : null;
  if (indentLeftPt !== null) format.indentLeftPt = indentLeftPt;

  return format;
}

function readRowHeight(trHeight: Element): RowHeight | null {
  const pt = twipsToPt(wAttr(trHeight, "val"));
  if (pt === null || pt <= 0) return null;
  // With no hRule, Word treats the height as a floor
  const rule = wAttr(trHeight, "hRule") === "exact" ? "exact" : "atLeast";
  return { rule, pt };
}

export function readRowFormat(trPr: Element | null): RowFormat | null {
  if (!trPr) return null;
  const trHeight = childByLocalName(trPr, "trHeight");
  const height = trHeight ? readRowHeight(trHeight) : null;
  const format: RowFormat = {};
  if (height) format.height = height;
  if (isOn(trPr, "tblHeader")) format.repeatHeader = true;
  if (isOn(trPr, "cantSplit")) format.cantSplit = true;
  return format;
}

const CELL_VERTICAL_ALIGN_BY_VAL: Record<string, CellVerticalAlign> = {
  top: "top",
  center: "center",
  bottom: "bottom",
};

/**
 * The padding of one cell: its own `w:tcMar` where it wrote one, and the table's cell margins on
 * the sides it did not. A side neither of them mentions is left absent, and the fallback on screen
 * is then the stylesheet's.
 */
function readCellPadding(
  tcPr: Element | null,
  margins: CellMargins
): Partial<CellFormat> {
  const padding = layerCellMargins(margins, readCellMarginsOf(tcPr, "tcMar"));
  const format: Partial<CellFormat> = {};
  if (padding.topPt !== null) format.paddingTopPt = padding.topPt;
  if (padding.rightPt !== null) format.paddingRightPt = padding.rightPt;
  if (padding.bottomPt !== null) format.paddingBottomPt = padding.bottomPt;
  if (padding.leftPt !== null) format.paddingLeftPt = padding.leftPt;
  return format;
}

/**
 * Cell formatting.
 * A side for which the cell has not written down its own border uses the line that side falls on,
 * which is the table's outer border or its inside line depending on where the cell sits.
 * The padding works the same way: what the cell wrote down wins, and the table's cell margins lie
 * underneath it. null if there is nothing at all to draw.
 */
export function readCellFormat(
  tcPr: Element | null,
  defaults: CellBorderDefaults,
  margins: CellMargins = NO_CELL_MARGINS
): CellFormat | null {
  const borders = tcPr ? childByLocalName(tcPr, "tcBorders") : null;
  const format: CellFormat = {};
  const top = borderSide(borders, "top") ?? defaults.top;
  if (top) format.borderTop = top;
  const bottom = borderSide(borders, "bottom") ?? defaults.bottom;
  if (bottom) format.borderBottom = bottom;
  const left =
    borderSide(borders, "left") ??
    borderSide(borders, "start") ??
    defaults.left;
  if (left) format.borderLeft = left;
  const right =
    borderSide(borders, "right") ??
    borderSide(borders, "end") ??
    defaults.right;
  if (right) format.borderRight = right;

  if (tcPr) {
    const background = shadingOf(tcPr);
    if (background) format.background = background;
    const verticalAlign =
      CELL_VERTICAL_ALIGN_BY_VAL[childValue(tcPr, "vAlign") ?? ""];
    if (verticalAlign) format.verticalAlign = verticalAlign;
  }
  // The table's margins reach a cell that wrote no formatting of its own at all
  const padded: CellFormat = { ...format, ...readCellPadding(tcPr, margins) };
  return Object.keys(padded).length > 0 ? padded : null;
}

/** The inside lines the table lays down, read from the table's own formatting XML */
export function insideBordersOf(tblPr: string | null): InsideBorders {
  const el = tblPr === null ? null : parsePropsXml(tblPr);
  return readInsideBorders(el);
}

/** The cell margins the table lays down, read from the table's own formatting XML */
export function cellMarginsOf(tblPr: string | null): CellMargins {
  const el = tblPr === null ? null : parsePropsXml(tblPr);
  return readCellMarginsOf(el, "tblCellMar");
}

/**
 * Reads a cell formatting fragment we operated on back into display values, along the same path
 * the import takes. The same fragment always yields the same display values.
 */
export function readCellProps(
  tcPr: string | null,
  defaults: CellBorderDefaults,
  margins: CellMargins = NO_CELL_MARGINS
): CellFormat | null {
  const el = tcPr === null ? null : parsePropsXml(tcPr);
  return readCellFormat(el, defaults, margins);
}

/** The formatting a cell holds. The original XML and the display values read out of it form a pair */
