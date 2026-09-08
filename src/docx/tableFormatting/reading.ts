/**
 * Reads table formatting for display. Table borders and margins are folded into cell values so CSS
 * collapsed-border resolution matches OOXML inheritance without changing exported XML.
 */

import {
  type BandSizes,
  type CellFormat,
  type CellMargins,
  type CellStyleBorders,
  type CellStyleFormat,
  type CellVerticalAlign,
  type InsideBorders,
  type RowFormat,
  type RowHeight,
  type RunFormat,
  TABLE_STYLE_CONDITIONS,
  type TableFormat,
  type TableStyleConditions,
  type TableStyleOverrideType,
  type TableWidth,
} from "../../model/format";
import { parsePropsXml } from "../../ooxml/props";
import {
  ST_DecimalNumber,
  ST_MeasurementOrPercent,
  ST_SignedTwipsMeasure,
  ST_TwipsMeasure,
} from "../../ooxml/simpleTypes";
import {
  ALIGN_BY_JC,
  type BorderLine,
  borderLineCss,
  borderSide,
  childValue,
  isOn,
  shadingOf,
  twipsToPt,
  wAttr,
} from "../../ooxml/units";
import { childByLocalName } from "../../ooxml/xml";
import type { ParagraphPlacement } from "../formatting/resolve";
import type { ParagraphFormatLayer } from "../formatting/tabStops";

export type {
  BandSizes,
  CellMargins,
  CellStyleFormat,
  InsideBorders,
  TableStyleConditions,
  TableStyleOverrideType,
} from "../../model/format";

/** A `pct` width counts in fiftieths of a percent, so 2500 and `50%` are the same width */
const FIFTIETHS_PER_PERCENT = 50;

/**
 * The width written down by `<w:tblW>` or `<w:tcW>`.
 * If we cannot make out what it means it is null, and such a width goes back out unchanged on export.
 *
 * `auto` and `nil` leave the width to the layout unless the value carries an explicit percentage.
 * A percentage may be written in fiftieths of a percent, as in `w:w="2500"`, or as `w:w="50%"`.
 * Both have to be gathered into the same unit, or the table collapses into a thin strip on screen.
 * Word lets an explicit `%` override `w:type` (MS-OI29500 §2.1.185(b)); see
 * spec/notes/simpleTypes.md. A universal measure under `pct` remains unreadable here.
 */
export function readTableWidth(
  parent: Element | null,
  name: "tblW" | "tcW"
): TableWidth | null {
  const el = parent ? childByLocalName(parent, name) : null;
  if (!el) return null;
  const type = wAttr(el, "type") ?? "dxa";
  const width = ST_MeasurementOrPercent.parse(wAttr(el, "w"));
  if (width?.kind === "percent") {
    const fiftieths = Math.round(width.value * FIFTIETHS_PER_PERCENT);
    return Number.isFinite(fiftieths) ? { type: "pct", fiftieths } : null;
  }
  if (type === "auto") return { type: "auto" };
  if (type === "nil") return { type: "nil" };

  if (width === null) return null;
  if (type === "dxa") {
    // A universal measure states an absolute width, which `universalMeasureToTwips` already gave
    return { type: "dxa", twips: width.value };
  }
  if (type === "pct") {
    if (width.kind === "twips") return null;
    return {
      type: "pct",
      fiftieths: width.value,
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
    const w = ST_TwipsMeasure.parse(wAttr(child, "w"));
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
    return type === "dxa"
      ? twipsToPt(ST_TwipsMeasure.parse(wAttr(el, "w")))
      : null;
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

/**
 * The lines a cell falls back on, one per side, already resolved for its spot in the grid.
 *
 * They are held as the markup records them rather than as CSS, because a cell that takes one of
 * them over as a border of its own writes it back into `w:tcBorders`.
 */
export interface CellBorderDefaults {
  top: BorderLine | null;
  bottom: BorderLine | null;
  left: BorderLine | null;
  right: BorderLine | null;
}

export const NO_BORDER_DEFAULTS: CellBorderDefaults = {
  top: null,
  bottom: null,
  left: null,
  right: null,
};

/** Which sides of a cell lie on the edge of a part of the table */
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
 * Everything one cell falls back on for what it did not write down itself, worked out from where
 * it sits in the grid by `cellDefaultsFor`. Everything that reads or edits a cell takes one of
 * these, so the lines a cell draws and the lines an edit measures against are always the same.
 */
export interface CellDefaults {
  borders: CellBorderDefaults;
  margins: CellMargins;
  background: string | null;
  verticalAlign: CellVerticalAlign | null;
  /** What the paragraphs inside the cell resolve their formatting against */
  placement: ParagraphPlacement;
}

export const NO_CELL_DEFAULTS: CellDefaults = {
  borders: NO_BORDER_DEFAULTS,
  margins: NO_CELL_MARGINS,
  background: null,
  verticalAlign: null,
  placement: { tableStyleId: null, conditions: [] },
};

export const NO_BAND_SIZES: BandSizes = { row: null, col: null };

/**
 * A band is made of whole rows or whole columns, so a count that is not a positive whole number
 * leaves the side unsaid and its bands are one row or one column wide (§17.7.6.5, §17.7.6.7).
 */
function bandSize(tblPr: Element, name: string): number | null {
  const count = ST_DecimalNumber.parse(childValue(tblPr, name));
  return count !== null && count > 0 ? count : null;
}

/** How many rows and columns one band of a table style is made of, read from the style's `w:tblPr` */
export function readBandSizes(tblPr: Element | null): BandSizes {
  if (!tblPr) return NO_BAND_SIZES;
  return {
    row: bandSize(tblPr, "tblStyleRowBandSize"),
    col: bandSize(tblPr, "tblStyleColBandSize"),
  };
}

export function layerBandSizes(style: BandSizes, over: BandSizes): BandSizes {
  return { row: over.row ?? style.row, col: over.col ?? style.col };
}

const NO_CELL_STYLE_BORDERS: CellStyleBorders = {
  top: null,
  bottom: null,
  left: null,
  right: null,
};

export const NO_CELL_STYLE_FORMAT: CellStyleFormat = {
  background: null,
  borders: NO_CELL_STYLE_BORDERS,
  inside: NO_INSIDE_BORDERS,
  margins: NO_CELL_MARGINS,
  verticalAlign: null,
};

export const NO_TABLE_STYLE_CONDITIONS: TableStyleConditions = {};

/**
 * The formatting one conditional format of a table style lays down (`w:tblStylePr`).
 * Its `w:tblPr` speaks for the region as a whole and its `w:tcPr` for the cells in it, so the
 * cell values are read with the table values underneath them.
 */
export interface ConditionalTableFormat {
  paragraph: ParagraphFormatLayer;
  run: RunFormat;
  table: TableFormat;
  cell: CellStyleFormat;
}

/** The lines a conditional format draws around its region: what its `w:tcBorders` writes, over what its `w:tblBorders` did */
function conditionBorders(
  tcBorders: Element | null,
  table: TableFormat
): CellStyleBorders {
  return {
    top: borderSide(tcBorders, "top") ?? table.borderTop ?? null,
    bottom: borderSide(tcBorders, "bottom") ?? table.borderBottom ?? null,
    left:
      borderSide(tcBorders, "left") ??
      borderSide(tcBorders, "start") ??
      table.borderLeft ??
      null,
    right:
      borderSide(tcBorders, "right") ??
      borderSide(tcBorders, "end") ??
      table.borderRight ??
      null,
  };
}

/**
 * What a conditional format lays down for the cells of its region.
 *
 * The `w:tcPr` of a conditional format writes the lines of the region rather than of every cell in
 * it: `insideH` and `insideV` are the lines between the cells it covers, which is how the header
 * row of a style with `insideV` set to `nil` is drawn as one unbroken band.
 */
export function readCellStyleFormat(
  tcPr: Element | null,
  tblPr: Element | null
): CellStyleFormat {
  const table = readTableFormat(tblPr) ?? {};
  const tcBorders = tcPr ? childByLocalName(tcPr, "tcBorders") : null;
  return {
    background: (tcPr ? shadingOf(tcPr) : null) ?? table.background ?? null,
    borders: conditionBorders(tcBorders, table),
    inside: layerInsideBorders(readInsideBorders(tblPr), {
      horizontal: borderSide(tcBorders, "insideH"),
      vertical: borderSide(tcBorders, "insideV"),
    }),
    margins: readCellMarginsOf(tcPr, "tcMar"),
    verticalAlign: tcPr ? cellVerticalAlignOf(tcPr) : null,
  };
}

function layerCellStyleBorders(
  base: CellStyleBorders,
  over: CellStyleBorders
): CellStyleBorders {
  return {
    top: over.top ?? base.top,
    bottom: over.bottom ?? base.bottom,
    left: over.left ?? base.left,
    right: over.right ?? base.right,
  };
}

/** Lays one conditional format over the one a style further up the `basedOn` chain wrote for the same part */
export function layerCellStyleFormat(
  base: CellStyleFormat,
  over: CellStyleFormat
): CellStyleFormat {
  return {
    background: over.background ?? base.background,
    borders: layerCellStyleBorders(base.borders, over.borders),
    inside: layerInsideBorders(base.inside, over.inside),
    margins: layerCellMargins(base.margins, over.margins),
    verticalAlign: over.verticalAlign ?? base.verticalAlign,
  };
}

/** The conditions a table style lays down for its cells alone, which is what a table carries for its cells to draw */
export function cellConditionsOf(
  conditions: Readonly<
    Partial<Record<TableStyleOverrideType, ConditionalTableFormat>>
  >
): TableStyleConditions {
  const cells: Partial<Record<TableStyleOverrideType, CellStyleFormat>> = {};
  for (const type of TABLE_STYLE_CONDITIONS) {
    const format = conditions[type];
    if (format) cells[type] = format.cell;
  }
  return cells;
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
  const indentLeftPt = tblInd
    ? twipsToPt(ST_SignedTwipsMeasure.parse(wAttr(tblInd, "w")))
    : null;
  if (indentLeftPt !== null) format.indentLeftPt = indentLeftPt;

  return format;
}

function readRowHeight(trHeight: Element): RowHeight | null {
  const pt = twipsToPt(ST_TwipsMeasure.parse(wAttr(trHeight, "val")));
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

/** Where `w:vAlign` places the content of a cell. null for a cell that says nothing, and for a value we do not draw */
function cellVerticalAlignOf(tcPr: Element): CellVerticalAlign | null {
  return CELL_VERTICAL_ALIGN_BY_VAL[childValue(tcPr, "vAlign") ?? ""] ?? null;
}

/**
 * The padding of one cell: its own `w:tcMar` where it wrote one, and the margins the table and its
 * style laid down on the sides it did not. A side neither of them mentions is left absent, and the
 * fallback on screen is then the stylesheet's.
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
 * which is the table's outer border, its inside line, or the line a conditional format of its
 * style draws there, depending on where the cell sits.
 * The fill, the vertical alignment and the padding work the same way: what the cell wrote down
 * wins, and what it falls back on lies underneath. null if there is nothing at all to draw.
 */
export function readCellFormat(
  tcPr: Element | null,
  defaults: CellDefaults = NO_CELL_DEFAULTS
): CellFormat | null {
  const borders = tcPr ? childByLocalName(tcPr, "tcBorders") : null;
  const lines = defaults.borders;
  const format: CellFormat = {};
  const top = borderSide(borders, "top") ?? borderLineCss(lines.top);
  if (top) format.borderTop = top;
  const bottom = borderSide(borders, "bottom") ?? borderLineCss(lines.bottom);
  if (bottom) format.borderBottom = bottom;
  const left =
    borderSide(borders, "left") ??
    borderSide(borders, "start") ??
    borderLineCss(lines.left);
  if (left) format.borderLeft = left;
  const right =
    borderSide(borders, "right") ??
    borderSide(borders, "end") ??
    borderLineCss(lines.right);
  if (right) format.borderRight = right;

  const background = (tcPr ? shadingOf(tcPr) : null) ?? defaults.background;
  if (background) format.background = background;
  const verticalAlign =
    (tcPr ? cellVerticalAlignOf(tcPr) : null) ?? defaults.verticalAlign;
  if (verticalAlign) format.verticalAlign = verticalAlign;
  // The margins reach a cell that wrote no formatting of its own at all
  const padded: CellFormat = {
    ...format,
    ...readCellPadding(tcPr, defaults.margins),
  };
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
  defaults: CellDefaults = NO_CELL_DEFAULTS
): CellFormat | null {
  const el = tcPr === null ? null : parsePropsXml(tcPr);
  return readCellFormat(el, defaults);
}

/** The formatting a cell holds. The original XML and the display values read out of it form a pair */
