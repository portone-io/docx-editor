import {
  type CellFormat,
  type DocumentDefaults,
  type HighlightName,
  type LineSpacing,
  NO_FILL,
  type ParagraphFormat,
  type RowFormat,
  type RunFormat,
  type TableFormat,
  type TableWidth,
  type UnderlineKind,
} from "../model/format";
import { editorCssVariables } from "./classNames";
import {
  DEFAULT_FONT_FALLBACKS,
  type FontFallbacks,
  withFontFallback,
} from "./fontStack";

/**
 * The value used on screen only, when nowhere in the document declares a default
 * font size.
 * It matches the value Word uses, and it never makes its way back into the document.
 */
export const FALLBACK_FONT_SIZE_PT = 10;

/**
 * What Word calls "one line" is not the font size but the line height the font
 * itself declares.
 * On screen that height is approximated as this multiple of the font size.
 * It is also the default for a document that declares no line spacing anywhere.
 *
 * Malgun Gothic, which dominates contracts, has a tall font-declared line box, so
 * Word really does render more loosely.
 * Drawing at 1.2 drew the complaint that it looked denser than the original, so it
 * was raised to 1.3.
 * The same value is also written as the `--docx-editor-line-height` default in
 * `editor.css`.
 */
export const SINGLE_LINE_RATIO = 1.3;

/**
 * Word highlighter (`w:highlight`) names and the colors they stand for.
 *
 * We write background colors with `w:shd`, but older documents recorded them under
 * these names.
 * Converting the names to colors lets such a document be drawn as the original was
 * and lets the toolbar palette compare against the current color.
 */
export const HIGHLIGHT_COLORS: Record<HighlightName, string> = {
  black: "#000000",
  blue: "#0000ff",
  cyan: "#00ffff",
  darkBlue: "#000080",
  darkCyan: "#008080",
  darkGray: "#808080",
  darkGreen: "#008000",
  darkMagenta: "#800080",
  darkRed: "#800000",
  darkYellow: "#808000",
  green: "#00ff00",
  lightGray: "#c0c0c0",
  magenta: "#ff00ff",
  red: "#ff0000",
  white: "#ffffff",
  yellow: "#ffff00",
};

/** Underline kinds mapped to CSS text-decoration-style. Any kind not listed is a straight line */
const UNDERLINE_STYLES: Partial<Record<UnderlineKind, string>> = {
  double: "double",
  dotted: "dotted",
  dottedHeavy: "dotted",
  dash: "dashed",
  dashedHeavy: "dashed",
  dashLong: "dashed",
  wave: "wavy",
  wavyHeavy: "wavy",
  wavyDouble: "wavy",
};

function pt(value: number): string {
  return `${value}pt`;
}

/** Trims the messy trailing decimals so the same input always yields the same value */
function round(value: number): number {
  return Math.round(value * 100) / 100;
}

function declarations(css: string[]): string | undefined {
  return css.length > 0 ? css.join(";") : undefined;
}

/**
 * The color a fill is painted with.
 * A fill of nothing is drawn as transparent rather than left undeclared, so that it wins over a
 * fill lying behind it.
 */
function fillColor(background: string): string {
  return background === NO_FILL ? "transparent" : background;
}

/** Turns line spacing into a single CSS line-height value. With none given, one line height is the default */
export function lineHeightValue(spacing: LineSpacing | null): string {
  if (!spacing) return `${SINGLE_LINE_RATIO}`;
  if (spacing.rule === "auto") {
    return `${round(spacing.lines * SINGLE_LINE_RATIO)}`;
  }
  if (spacing.rule === "exact") return pt(spacing.pt);
  // atLeast is the larger of the pinned height and one line height
  return `max(${pt(spacing.pt)},${SINGLE_LINE_RATIO}em)`;
}

function textDecoration(format: RunFormat): string[] {
  const lines: string[] = [];
  if (format.underline) lines.push("underline");
  if (format.strike) lines.push("line-through");
  if (lines.length === 0) return [];

  const css = [`text-decoration-line:${lines.join(" ")}`];
  const style = format.underline
    ? UNDERLINE_STYLES[format.underline]
    : undefined;
  if (style) css.push(`text-decoration-style:${style}`);
  return css;
}

/** The character declarations that carry down to the text inside the element that wrote them */
function inheritedRunCss(
  format: RunFormat,
  fontFallbacks: FontFallbacks
): string[] {
  const css: string[] = [];
  if (format.bold) css.push("font-weight:bold");
  if (format.italic) css.push("font-style:italic");
  if (format.smallCaps) css.push("font-variant:small-caps");
  css.push(...textDecoration(format));
  if (format.fontSizePt !== undefined)
    css.push(`font-size:${pt(format.fontSizePt)}`);
  if (format.fontFamily)
    css.push(
      `font-family:${withFontFallback(format.fontFamily, fontFallbacks)}`
    );
  if (format.color) css.push(`color:${format.color}`);
  return css;
}

/** The fill painted behind the characters and where they sit on the line. Neither carries down */
function ownRunCss(format: RunFormat): string[] {
  const css: string[] = [];
  const highlight = format.highlight
    ? HIGHLIGHT_COLORS[format.highlight]
    : undefined;
  if (highlight) css.push(`background-color:${highlight}`);
  else if (format.background)
    css.push(`background-color:${fillColor(format.background)}`);
  if (format.verticalAlign === "superscript") css.push("vertical-align:super");
  if (format.verticalAlign === "subscript") css.push("vertical-align:sub");
  return css;
}

export function runStyle(
  format: RunFormat | null,
  fontFallbacks: FontFallbacks = DEFAULT_FONT_FALLBACKS
): string | undefined {
  if (!format) return undefined;
  return declarations([
    ...inheritedRunCss(format, fontFallbacks),
    ...ownRunCss(format),
  ]);
}

/**
 * The CSS a paragraph is drawn with: its own formatting, and the character formatting the style it
 * wears lays down for the text inside it (`styleRun`).
 *
 * Text typed in the editor carries no run of its own, so inheriting from the paragraph is the only
 * way it can be drawn in the style's character formatting. Text that does carry a run declares
 * those same values on itself, and an element's own declaration beats an inherited one, so nothing
 * a run wrote down is disturbed.
 * Only the declarations that carry down are written. A fill paints a box instead of passing down,
 * so on a paragraph it would cover the whole line, and the paragraph's own shading along with it.
 * A style that fills behind its characters is therefore drawn from the run each piece of text
 * carries, and not from here.
 * The fonts stood in for the ones the style names are the built-in ones, because unlike a run
 * (`editor/runMarkView`) a paragraph is drawn by the schema alone, which cannot see which editor it
 * is drawing for.
 */
export function paragraphStyle(
  format: ParagraphFormat | null,
  styleRun: RunFormat | null = null
): string | undefined {
  const inherited = styleRun
    ? inheritedRunCss(styleRun, DEFAULT_FONT_FALLBACKS)
    : [];
  if (!format) return declarations(inherited);
  const css: string[] = [...inherited];
  if (format.align) css.push(`text-align:${format.align}`);
  if (format.direction) css.push(`direction:${format.direction}`);
  if (format.indentLeftPt !== undefined)
    css.push(`margin-left:${pt(format.indentLeftPt)}`);
  if (format.indentRightPt !== undefined)
    css.push(`margin-right:${pt(format.indentRightPt)}`);
  if (format.indentStartPt !== undefined)
    css.push(`margin-inline-start:${pt(format.indentStartPt)}`);
  if (format.indentEndPt !== undefined)
    css.push(`margin-inline-end:${pt(format.indentEndPt)}`);
  if (format.textIndentPt !== undefined)
    css.push(`text-indent:${pt(format.textIndentPt)}`);
  if (format.spaceBeforePt !== undefined)
    css.push(`margin-top:${pt(format.spaceBeforePt)}`);
  if (format.spaceAfterPt !== undefined)
    css.push(`margin-bottom:${pt(format.spaceAfterPt)}`);
  if (format.lineSpacing)
    css.push(`line-height:${lineHeightValue(format.lineSpacing)}`);
  if (format.borderTop) css.push(`border-top:${format.borderTop}`);
  if (format.borderBottom) css.push(`border-bottom:${format.borderBottom}`);
  if (format.borderLeft) css.push(`border-left:${format.borderLeft}`);
  if (format.borderRight) css.push(`border-right:${format.borderRight}`);
  if (format.background)
    css.push(`background-color:${fillColor(format.background)}`);
  return declarations(css);
}

/**
 * Table grid widths are recorded in twips (dxa).
 * One pt is 20 twips and one pt on screen is 96/72 pixels, so dividing by 15 gives
 * pixels.
 */
const DXA_PER_PX = 15;

export function columnWidthPx(dxa: number): number {
  return round(dxa / DXA_PER_PX);
}

/** A `pct` width is measured in fiftieths of a percent. The other kinds ask for no width at all */
function widthCss(width: TableWidth | null): string[] {
  if (width?.type === "dxa") return [`width:${columnWidthPx(width.twips)}px`];
  if (width?.type === "pct") return [`width:${round(width.fiftieths / 50)}%`];
  return [];
}

function borderDeclarations(format: {
  borderTop?: string;
  borderBottom?: string;
  borderLeft?: string;
  borderRight?: string;
}): string[] {
  const css: string[] = [];
  if (format.borderTop) css.push(`border-top:${format.borderTop}`);
  if (format.borderBottom) css.push(`border-bottom:${format.borderBottom}`);
  if (format.borderLeft) css.push(`border-left:${format.borderLeft}`);
  if (format.borderRight) css.push(`border-right:${format.borderRight}`);
  return css;
}

const TABLE_ALIGN_MARGIN: Record<string, string> = {
  center: "margin-left:auto;margin-right:auto",
  right: "margin-left:auto;margin-right:0",
};

/**
 * The table itself draws no line at all: every line of a table, the ones around the outside included,
 * is drawn by its cells. A border on the table element would win over a cell that says to draw
 * nothing, because CSS resolves a collapsed border in favour of the solid line.
 */
export function tableStyle(
  format: TableFormat | null,
  width: TableWidth | null
): string | undefined {
  const css: string[] = [...widthCss(width)];
  if (format) {
    if (format.background)
      css.push(`background-color:${fillColor(format.background)}`);
    // A table standing at the margin (an indent of 0) is still free to be aligned
    const align = format.align ? TABLE_ALIGN_MARGIN[format.align] : undefined;
    if (align && !format.indentLeftPt) css.push(align);
    else if (format.indentLeftPt !== undefined)
      css.push(`margin-left:${pt(format.indentLeftPt)}`);
  }
  return declarations(css);
}

/** CSS treats a row's height as a floor whichever rule the document wrote, so both are drawn alike */
export function rowStyle(format: RowFormat | null): string | undefined {
  const height = format?.height;
  return height ? `height:${pt(height.pt)}` : undefined;
}

const CELL_VERTICAL_ALIGN: Record<string, string> = {
  top: "top",
  center: "middle",
  bottom: "bottom",
};

export function cellStyle(format: CellFormat | null): string | undefined {
  if (!format) return undefined;
  const css: string[] = [...borderDeclarations(format)];
  if (format.background)
    css.push(`background-color:${fillColor(format.background)}`);
  const verticalAlign = format.verticalAlign
    ? CELL_VERTICAL_ALIGN[format.verticalAlign]
    : undefined;
  if (verticalAlign) css.push(`vertical-align:${verticalAlign}`);
  if (format.paddingTopPt !== undefined)
    css.push(`padding-top:${pt(format.paddingTopPt)}`);
  if (format.paddingRightPt !== undefined)
    css.push(`padding-right:${pt(format.paddingRightPt)}`);
  if (format.paddingBottomPt !== undefined)
    css.push(`padding-bottom:${pt(format.paddingBottomPt)}`);
  if (format.paddingLeftPt !== undefined)
    css.push(`padding-left:${pt(format.paddingLeftPt)}`);
  return declarations(css);
}

/**
 * Turns the document defaults into CSS variables that apply across the whole paper.
 * A font is always settled on, even when the document declares none, so that it is
 * not drawn in a different font in every environment.
 */
export function documentDefaultsStyle(
  defaults: DocumentDefaults,
  fontFallbacks: FontFallbacks = DEFAULT_FONT_FALLBACKS
): string {
  const fontFamily = defaults.fontFamily
    ? withFontFallback(defaults.fontFamily, fontFallbacks)
    : fontFallbacks.defaultStack;
  return [
    `${editorCssVariables.fontSize}:${pt(defaults.fontSizePt ?? FALLBACK_FONT_SIZE_PT)}`,
    `${editorCssVariables.lineHeight}:${lineHeightValue(defaults.lineSpacing)}`,
    `${editorCssVariables.fontFamily}:${fontFamily}`,
  ].join(";");
}
