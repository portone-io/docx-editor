/**
 * The formatting values to render on screen, and the checks that they have the shape
 * they should.
 *
 * These values are produced by reading the original formatting XML (`w:pPr`, `w:rPr`,
 * `w:tblPr`, ...) and are only ever used for rendering. What goes back into the document
 * is always the original XML string.
 *
 * A value enters as a node attribute, leaves for the DOM as `data-fmt`, and comes back
 * in again. So that a value tampered with along that path cannot leak into a CSS
 * declaration, the checks that let through only the shapes the read functions produce
 * live here alongside them.
 */

import type { TabStop } from "./tabStops";

const ALIGNS = ["left", "center", "right", "justify"] as const;
export type ParagraphAlign = (typeof ALIGNS)[number];

const VERTICAL_ALIGNS = ["superscript", "subscript"] as const;
export type VerticalAlign = (typeof VERTICAL_ALIGNS)[number];

const CELL_VERTICAL_ALIGNS = ["top", "center", "bottom"] as const;
export type CellVerticalAlign = (typeof CELL_VERTICAL_ALIGNS)[number];

/**
 * The fill of something whose document says to paint nothing (`w:shd w:fill="auto"`).
 * A fill nobody mentioned is absent instead, so a style's fill comes back only in that case.
 * The lines use the same word for the same reason.
 */
export const NO_FILL = "none";

/**
 * The underline kinds OOXML defines (`ST_Underline`). "No underline" (`none`) is not
 * included
 */
const UNDERLINES = [
  "single",
  "words",
  "double",
  "thick",
  "dotted",
  "dottedHeavy",
  "dash",
  "dashedHeavy",
  "dashLong",
  "dashLongHeavy",
  "dotDash",
  "dashDotHeavy",
  "dotDotDash",
  "dashDotDotHeavy",
  "wave",
  "wavyHeavy",
  "wavyDouble",
] as const;
export type UnderlineKind = (typeof UNDERLINES)[number];

/**
 * The highlight names OOXML defines (`ST_HighlightColor`). "No highlight" (`none`) is
 * not included
 */
const HIGHLIGHTS = [
  "black",
  "blue",
  "cyan",
  "darkBlue",
  "darkCyan",
  "darkGray",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "green",
  "lightGray",
  "magenta",
  "red",
  "white",
  "yellow",
] as const;
export type HighlightName = (typeof HIGHLIGHTS)[number];

export function isUnderlineKind(value: unknown): value is UnderlineKind {
  return UNDERLINES.some((kind) => kind === value);
}

export function isHighlightName(value: unknown): value is HighlightName {
  return HIGHLIGHTS.some((name) => name === value);
}

/**
 * The width a table or a cell has recorded.
 * `auto` and `nil` leave the width up to the layout, so they carry no number at all.
 */
export type TableWidth =
  | { type: "dxa"; twips: number }
  | { type: "pct"; fiftieths: number }
  | { type: "auto" }
  | { type: "nil" };

export type TableWidthType = TableWidth["type"];

/** The number a width records, in the unit its type names */
export function widthNumber(width: TableWidth): number | null {
  if (width.type === "dxa") return width.twips;
  if (width.type === "pct") return width.fiftieths;
  return null;
}

export function withWidthNumber(width: TableWidth, value: number): TableWidth {
  if (width.type === "dxa") return { type: "dxa", twips: value };
  if (width.type === "pct") return { type: "pct", fiftieths: value };
  return width;
}

/**
 * Line spacing.
 * `auto` states a multiple of one line's height; the others pin the height down in
 * points.
 */
export type LineSpacing =
  | { rule: "auto"; lines: number }
  | { rule: "exact"; pt: number }
  | { rule: "atLeast"; pt: number };

/**
 * Which list this paragraph belongs to and which of its levels it sits at. The number
 * itself is computed at render time
 */
export interface NumberingRef {
  numId: number;
  ilvl: number;
}

/** The paragraph formatting to render on screen. Lengths are in points */
export interface ParagraphFormat {
  align?: ParagraphAlign;
  /** The paragraph's logical text direction (`w:bidi`). */
  direction?: "ltr" | "rtl";
  indentLeftPt?: number;
  indentRightPt?: number;
  /** Logical leading and trailing indents from Strict `w:start` and `w:end`. */
  indentStartPt?: number;
  indentEndPt?: number;
  /** Positive when only the first line is indented further, negative for a hanging indent */
  textIndentPt?: number;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacing?: LineSpacing;
  numbering?: NumberingRef;
  /** Effective custom tab stops after the paragraph-property hierarchy is resolved. */
  tabStops?: readonly TabStop[];
  /**
   * Set when the document records that a new page starts at this paragraph. Read by the
   * page boundary guides
   */
  pageBreakBefore?: true;
  /** A CSS border in the form `1pt solid #000000` */
  borderTop?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderRight?: string;
  background?: string;
}

/** The character formatting to render on screen */
export interface RunFormat {
  bold?: true;
  italic?: true;
  /** The underline kind. Absent entirely when there is no underline */
  underline?: UnderlineKind;
  strike?: true;
  smallCaps?: true;
  fontSizePt?: number;
  /** A CSS font name list in the form `"맑은 고딕","Malgun Gothic"` */
  fontFamily?: string;
  color?: string;
  highlight?: HighlightName;
  background?: string;
  verticalAlign?: VerticalAlign;
  /**
   * The language the run records (`w:lang`), drawn as the `lang` attribute of its span.
   * Han characters are written differently in Japanese, in simplified Chinese and in
   * traditional Chinese, and this is the only thing that tells the browser which of
   * those shapes to draw. It is read for display only and never written back.
   */
  lang?: string;
}

/**
 * The table formatting to render on screen. The inner lines between cells are folded
 * into the cell formatting
 */
export interface TableFormat {
  borderTop?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderRight?: string;
  background?: string;
  align?: ParagraphAlign;
  indentLeftPt?: number;
}

/**
 * The lines that divide the inside of a table (`insideH`, `insideV`).
 * They are held by the table but drawn by the cells, so they become the defaults for cell
 * formatting. A side with nothing written down is null.
 */
export interface InsideBorders {
  horizontal: string | null;
  vertical: string | null;
}

/**
 * The cell margins a table or its style lays down, in points, one per side.
 * They are held by the table but drawn by the cells, so they become the padding a cell
 * that wrote none of its own falls back on. A side with nothing written down is null.
 */
export interface CellMargins {
  topPt: number | null;
  rightPt: number | null;
  bottomPt: number | null;
  leftPt: number | null;
}

/**
 * How high a row stands, in points.
 * `atLeast` is a floor the row grows past when its text needs the room, and `exact` pins
 * the height down.
 */
export type RowHeight =
  | { rule: "atLeast"; pt: number }
  | { rule: "exact"; pt: number };

/** The formatting of a single table row to render on screen */
export interface RowFormat {
  /** What `w:trHeight` writes down */
  height?: RowHeight;
  /** Whether this row is repeated at the top of a continued table */
  repeatHeader?: boolean;
  /** Whether this row must stay whole when the table continues on another page */
  cantSplit?: boolean;
}

/** The formatting of a single table cell to render on screen */
export interface CellFormat {
  borderTop?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderRight?: string;
  background?: string;
  verticalAlign?: CellVerticalAlign;
  paddingTopPt?: number;
  paddingRightPt?: number;
  paddingBottomPt?: number;
  paddingLeftPt?: number;
}

/**
 * The defaults that apply to the whole document. Null when absent; the on-screen
 * fallback is decided by the presentation layer
 */
export interface DocumentDefaults {
  fontSizePt: number | null;
  fontFamily: string | null;
  lineSpacing: LineSpacing | null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * The shapes a display value must have to reach a CSS declaration: only what the read
 * functions (`docx/formatting`, `docx/tableFormatting`) produce is let through.
 */
const HEX_COLOR = /^#[0-9a-fA-F]{6}$/;

/** The form `0.5pt solid #000000`. `none` means "do not draw this side" */
const BORDER_CSS =
  /^(none|\d+(\.\d{1,2})?pt (solid|double|dotted|dashed) #[0-9a-fA-F]{6})$/;

/** A fill: a color, or `none` where the document says to paint nothing (`NO_FILL`) */
const FILL = /^(none|#[0-9a-fA-F]{6})$/;

/**
 * The form `"맑은 고딕","Malgun Gothic"`. A name containing a quote or a semicolon would
 * break the declaration
 */
const FONT_FAMILY = /^"[^"';]+"(,"[^"';]+")*$/;

/**
 * A language tag, the form `ja-JP` or `zh-Hant-TW`.
 * It leaves for the screen as an attribute, so only the shape a tag has is let through.
 */
const LANGUAGE_TAG = /^[A-Za-z]{1,8}(-[A-Za-z0-9]{1,8}){0,4}$/;

function matching(pattern: RegExp, value: unknown): string | undefined {
  return typeof value === "string" && pattern.test(value) ? value : undefined;
}

function isAlign(value: unknown): value is ParagraphAlign {
  return ALIGNS.some((align) => align === value);
}

function isVerticalAlign(value: unknown): value is VerticalAlign {
  return VERTICAL_ALIGNS.some((align) => align === value);
}

function toLineSpacing(value: unknown): LineSpacing | null {
  if (!isRecord(value)) return null;
  if (value.rule === "auto" && typeof value.lines === "number") {
    return { rule: "auto", lines: value.lines };
  }
  if (typeof value.pt !== "number") return null;
  if (value.rule === "exact") return { rule: "exact", pt: value.pt };
  if (value.rule === "atLeast") return { rule: "atLeast", pt: value.pt };
  return null;
}

function toNumberingRef(value: unknown): NumberingRef | null {
  if (!isRecord(value)) return null;
  const { numId, ilvl } = value;
  if (typeof numId !== "number" || typeof ilvl !== "number") return null;
  return { numId, ilvl };
}

const TAB_ALIGNMENTS = [
  "start",
  "center",
  "end",
  "decimal",
  "num",
  "bar",
] as const;

const TAB_LEADERS = [
  "none",
  "dot",
  "hyphen",
  "underscore",
  "heavy",
  "middleDot",
] as const;

function toTabStops(value: unknown): TabStop[] | null {
  if (!Array.isArray(value)) return null;
  const stops: TabStop[] = [];
  for (const entry of value) {
    if (
      !isRecord(entry) ||
      typeof entry.positionPt !== "number" ||
      !Number.isFinite(entry.positionPt)
    ) {
      continue;
    }
    const align = TAB_ALIGNMENTS.find((candidate) => candidate === entry.align);
    if (!align) continue;
    const leader = TAB_LEADERS.find((candidate) => candidate === entry.leader);
    stops.push({
      positionPt: entry.positionPt,
      align,
      ...(leader === undefined ? {} : { leader }),
    });
  }
  return stops.length > 0 ? stops : null;
}

export function toParagraphFormat(value: unknown): ParagraphFormat | null {
  if (!isRecord(value)) return null;
  const format: ParagraphFormat = {};
  if (isAlign(value.align)) format.align = value.align;
  if (value.direction === "ltr" || value.direction === "rtl")
    format.direction = value.direction;
  if (typeof value.indentLeftPt === "number")
    format.indentLeftPt = value.indentLeftPt;
  if (typeof value.indentRightPt === "number")
    format.indentRightPt = value.indentRightPt;
  if (typeof value.indentStartPt === "number")
    format.indentStartPt = value.indentStartPt;
  if (typeof value.indentEndPt === "number")
    format.indentEndPt = value.indentEndPt;
  if (typeof value.textIndentPt === "number")
    format.textIndentPt = value.textIndentPt;
  if (typeof value.spaceBeforePt === "number")
    format.spaceBeforePt = value.spaceBeforePt;
  if (typeof value.spaceAfterPt === "number")
    format.spaceAfterPt = value.spaceAfterPt;
  const lineSpacing = toLineSpacing(value.lineSpacing);
  if (lineSpacing) format.lineSpacing = lineSpacing;
  const numbering = toNumberingRef(value.numbering);
  if (numbering) format.numbering = numbering;
  const tabStops = toTabStops(value.tabStops);
  if (tabStops) format.tabStops = tabStops;
  if (value.pageBreakBefore === true) format.pageBreakBefore = true;
  copyBorders(value, format);
  const background = matching(FILL, value.background);
  if (background) format.background = background;
  return format;
}

export function toRunFormat(value: unknown): RunFormat | null {
  if (!isRecord(value)) return null;
  const format: RunFormat = {};
  if (value.bold === true) format.bold = true;
  if (value.italic === true) format.italic = true;
  if (value.strike === true) format.strike = true;
  if (value.smallCaps === true) format.smallCaps = true;
  if (isUnderlineKind(value.underline)) format.underline = value.underline;
  if (typeof value.fontSizePt === "number")
    format.fontSizePt = value.fontSizePt;
  const fontFamily = matching(FONT_FAMILY, value.fontFamily);
  if (fontFamily) format.fontFamily = fontFamily;
  const color = matching(HEX_COLOR, value.color);
  if (color) format.color = color;
  if (isHighlightName(value.highlight)) format.highlight = value.highlight;
  const background = matching(FILL, value.background);
  if (background) format.background = background;
  if (isVerticalAlign(value.verticalAlign))
    format.verticalAlign = value.verticalAlign;
  const lang = matching(LANGUAGE_TAG, value.lang);
  if (lang) format.lang = lang;
  return format;
}

function isCellVerticalAlign(value: unknown): value is CellVerticalAlign {
  return CELL_VERTICAL_ALIGNS.some((align) => align === value);
}

export function toTableWidth(value: unknown): TableWidth | null {
  if (!isRecord(value)) return null;
  if (value.type === "auto") return { type: "auto" };
  if (value.type === "nil") return { type: "nil" };
  if (value.type === "dxa" && typeof value.twips === "number")
    return { type: "dxa", twips: value.twips };
  if (value.type === "pct" && typeof value.fiftieths === "number")
    return { type: "pct", fiftieths: value.fiftieths };
  return null;
}

/** The sequence of column widths (dxa) that `w:tblGrid` defines */
export function toGridCols(value: unknown): number[] {
  if (!Array.isArray(value)) return [];
  return value.filter((entry): entry is number => typeof entry === "number");
}

/** The cell widths (px) `prosemirror-tables` uses. Our documents usually leave it empty */
export function toColWidth(value: unknown): number[] | null {
  if (!Array.isArray(value)) return null;
  const widths = value.filter(
    (entry): entry is number => typeof entry === "number"
  );
  return widths.length === value.length && widths.length > 0 ? widths : null;
}

function copyBorders(
  value: Record<string, unknown>,
  format: {
    borderTop?: string;
    borderBottom?: string;
    borderLeft?: string;
    borderRight?: string;
  }
): void {
  const top = matching(BORDER_CSS, value.borderTop);
  if (top) format.borderTop = top;
  const bottom = matching(BORDER_CSS, value.borderBottom);
  if (bottom) format.borderBottom = bottom;
  const left = matching(BORDER_CSS, value.borderLeft);
  if (left) format.borderLeft = left;
  const right = matching(BORDER_CSS, value.borderRight);
  if (right) format.borderRight = right;
}

export function toTableFormat(value: unknown): TableFormat | null {
  if (!isRecord(value)) return null;
  const format: TableFormat = {};
  copyBorders(value, format);
  const background = matching(FILL, value.background);
  if (background) format.background = background;
  if (isAlign(value.align)) format.align = value.align;
  if (typeof value.indentLeftPt === "number")
    format.indentLeftPt = value.indentLeftPt;
  return format;
}

/**
 * The inside lines a table style laid down, as they come back in from `data-style-inside`.
 * A table with no line on either side is the same as the attribute not being there at all,
 * so it is null.
 */
export function toInsideBorders(value: unknown): InsideBorders | null {
  if (!isRecord(value)) return null;
  const horizontal = matching(BORDER_CSS, value.horizontal) ?? null;
  const vertical = matching(BORDER_CSS, value.vertical) ?? null;
  if (horizontal === null && vertical === null) return null;
  return { horizontal, vertical };
}

function marginSide(value: unknown): number | null {
  return typeof value === "number" ? value : null;
}

/**
 * The cell margins a table style laid down, as they come back in from `data-style-margins`.
 * A table with a margin on none of its four sides is the same as the attribute not being
 * there at all, so it is null.
 */
export function toCellMargins(value: unknown): CellMargins | null {
  if (!isRecord(value)) return null;
  const margins: CellMargins = {
    topPt: marginSide(value.topPt),
    rightPt: marginSide(value.rightPt),
    bottomPt: marginSide(value.bottomPt),
    leftPt: marginSide(value.leftPt),
  };
  const sides = [
    margins.topPt,
    margins.rightPt,
    margins.bottomPt,
    margins.leftPt,
  ];
  return sides.every((side) => side === null) ? null : margins;
}

function toRowHeight(value: unknown): RowHeight | null {
  if (!isRecord(value) || typeof value.pt !== "number") return null;
  if (value.rule === "exact") return { rule: "exact", pt: value.pt };
  if (value.rule === "atLeast") return { rule: "atLeast", pt: value.pt };
  return null;
}

export function toRowFormat(value: unknown): RowFormat | null {
  if (!isRecord(value)) return null;
  const format: RowFormat = {};
  const height = toRowHeight(value.height);
  if (height) format.height = height;
  if (value.repeatHeader === true) format.repeatHeader = true;
  if (value.cantSplit === true) format.cantSplit = true;
  return format;
}

export function toCellFormat(value: unknown): CellFormat | null {
  if (!isRecord(value)) return null;
  const format: CellFormat = {};
  copyBorders(value, format);
  const background = matching(FILL, value.background);
  if (background) format.background = background;
  if (isCellVerticalAlign(value.verticalAlign))
    format.verticalAlign = value.verticalAlign;
  if (typeof value.paddingTopPt === "number")
    format.paddingTopPt = value.paddingTopPt;
  if (typeof value.paddingRightPt === "number")
    format.paddingRightPt = value.paddingRightPt;
  if (typeof value.paddingBottomPt === "number")
    format.paddingBottomPt = value.paddingBottomPt;
  if (typeof value.paddingLeftPt === "number")
    format.paddingLeftPt = value.paddingLeftPt;
  return format;
}

/**
 * A merge count (colspan, rowspan).
 * Because it feeds grid arithmetic and loop steps, only positive integers are let
 * through. The fallback of 1 means "not merged".
 */
export function spanCount(value: unknown): number {
  return typeof value === "number" && Number.isInteger(value) && value > 0
    ? value
    : 1;
}
