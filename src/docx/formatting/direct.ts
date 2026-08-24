/**
 * Reads direct formatting XML into display values.
 *
 * The values produced here are used for display only.
 * What goes back into the document is always the original XML string, so no matter what is read
 * here, export is undisturbed.
 */

import {
  type DocumentDefaults,
  isHighlightName,
  isUnderlineKind,
  type LineSpacing,
  type ParagraphFormat,
  type RunFormat,
  type VerticalAlign,
} from "../../model/format";
import { readTabStopDirectives } from "../../ooxml/tabStops";
import {
  ALIGN_BY_JC,
  borderSide,
  childValue,
  halfPointsToPt,
  isOn,
  round,
  shadingOf,
  toHexColor,
  toNumber,
  twipsToPt,
  wAttr,
} from "../../ooxml/units";
import { childByLocalName } from "../../ooxml/xml";
import {
  NO_THEME_FONTS,
  THEME_ATTRS,
  type ThemeFonts,
  themeFontName,
} from "../theme";
import type { ParagraphFormatLayer } from "./tabStops";

const VERTICAL_ALIGN_BY_VAL: Record<string, VerticalAlign> = {
  superscript: "superscript",
  subscript: "subscript",
};

/**
 * The font names held in a display value.
 * Splits the CSS name list `fontFamilyOf` produced back into individual names.
 */
export function fontNamesOf(cssNames: string | undefined | null): string[] {
  if (!cssNames) return [];
  return cssNames
    .split(",")
    .map((name) => name.trim().replace(/^"|"$/g, ""))
    .filter((name) => name.length > 0);
}

/** The order the slots are read in. The Latin one leads, since that is the font a mixed run mostly shows */
const FONT_SLOTS: readonly string[] = ["ascii", "eastAsia", "hAnsi", "cs"];

/**
 * The font one slot asks for: the name it wrote down, or else the font the theme
 * reference beside it stands for
 */
function slotFontName(
  rFonts: Element,
  slot: string,
  themeFonts: ThemeFonts
): string | null {
  const written = wAttr(rFonts, slot);
  if (written !== null) return written;
  for (const attr of THEME_ATTRS[slot] ?? []) {
    const resolved = themeFontName(themeFonts, wAttr(rFonts, attr));
    if (resolved !== null) return resolved;
  }
  return null;
}

/** Builds the list to use on screen out of the several font names */
function fontFamilyOf(rFonts: Element, themeFonts: ThemeFonts): string | null {
  const names = FONT_SLOTS.map((slot) => slotFontName(rFonts, slot, themeFonts))
    .filter((name): name is string => name !== null)
    // A name holding a quote or a semicolon would break the CSS declaration it is written
    // into, and a trailing backslash would escape the quote we wrap it in, so we drop it
    .filter((name) => name.length > 0 && !/["';\\]/.test(name));
  const unique = Array.from(new Set(names));
  return unique.length > 0 ? unique.map((name) => `"${name}"`).join(",") : null;
}

/**
 * The language a run records (`w:lang`).
 *
 * `w:eastAsia` is the one that decides which shape of a Han character is drawn, so where a
 * run states one it wins over the `w:val` that holds the Latin language.
 * Nothing here reaches the file: `w:lang` goes back out inside the original rPr text.
 */
function langOf(rPr: Element): string | null {
  const lang = childByLocalName(rPr, "lang");
  if (!lang) return null;
  const eastAsia = wAttr(lang, "eastAsia");
  if (eastAsia !== null && eastAsia.length > 0) return eastAsia;
  const value = wAttr(lang, "val");
  return value !== null && value.length > 0 ? value : null;
}

/** Paragraph borders carry only the sides actually drawn. A side pinned down as not drawn is the same as none at all */
function drawnBorder(pBdr: Element, side: string): string | null {
  const css = borderSide(pBdr, side);
  return css === "none" ? null : css;
}

function readBorders(pPr: Element): Partial<ParagraphFormat> {
  const pBdr = childByLocalName(pPr, "pBdr");
  if (!pBdr) return {};
  const format: Partial<ParagraphFormat> = {};
  const top = drawnBorder(pBdr, "top");
  if (top) format.borderTop = top;
  const bottom = drawnBorder(pBdr, "bottom");
  if (bottom) format.borderBottom = bottom;
  const left = drawnBorder(pBdr, "left");
  if (left) format.borderLeft = left;
  const right = drawnBorder(pBdr, "right");
  if (right) format.borderRight = right;
  return format;
}

function readIndent(pPr: Element): Partial<ParagraphFormat> {
  const ind = childByLocalName(pPr, "ind");
  if (!ind) return {};
  const format: Partial<ParagraphFormat> = {};
  const start = twipsToPt(wAttr(ind, "start") ?? wAttr(ind, "left"));
  if (start !== null) format.indentStartPt = start;
  const end = twipsToPt(wAttr(ind, "end") ?? wAttr(ind, "right"));
  if (end !== null) format.indentEndPt = end;
  // A hanging indent overrides the first-line indent (an OOXML rule)
  const hanging = twipsToPt(wAttr(ind, "hanging"));
  const firstLine = twipsToPt(wAttr(ind, "firstLine"));
  if (hanging !== null) format.textIndentPt = -hanging;
  else if (firstLine !== null) format.textIndentPt = firstLine;
  return format;
}

/**
 * The line spacing from `w:spacing`.
 * auto means a multiple in 240ths, and the rest are heights pinned down in twips.
 */
export function readLineSpacing(pPr: Element): LineSpacing | null {
  const spacing = childByLocalName(pPr, "spacing");
  if (!spacing) return null;
  const line = toNumber(wAttr(spacing, "line"));
  if (line === null) return null;
  const rule = wAttr(spacing, "lineRule");
  if (rule === "exact") return { rule: "exact", pt: round(line / 20) };
  if (rule === "atLeast") return { rule: "atLeast", pt: round(line / 20) };
  return { rule: "auto", lines: round(line / 240) };
}

function readSpacing(pPr: Element): Partial<ParagraphFormat> {
  const spacing = childByLocalName(pPr, "spacing");
  if (!spacing) return {};
  const format: Partial<ParagraphFormat> = {};
  const before = twipsToPt(wAttr(spacing, "before"));
  if (before !== null) format.spaceBeforePt = before;
  const after = twipsToPt(wAttr(spacing, "after"));
  if (after !== null) format.spaceAfterPt = after;
  const lineSpacing = readLineSpacing(pPr);
  if (lineSpacing) format.lineSpacing = lineSpacing;
  return format;
}

/** The list slot `w:numPr` points at. A numId of 0 means "not a list" */
function readNumbering(pPr: Element): Partial<ParagraphFormatLayer> {
  const numPr = childByLocalName(pPr, "numPr");
  if (!numPr) return {};
  const numId = toNumber(childValue(numPr, "numId"));
  if (numId === 0) return { numbering: null };
  if (numId === null || numId < 0) return {};
  return {
    numbering: {
      numId: Math.trunc(numId),
      ilvl: Math.trunc(toNumber(childValue(numPr, "ilvl")) ?? 0),
    },
  };
}

export function readParagraphFormat(
  pPr: Element | null
): ParagraphFormatLayer | null {
  if (!pPr) return null;
  const format: ParagraphFormatLayer = {
    ...readIndent(pPr),
    ...readSpacing(pPr),
    ...readNumbering(pPr),
    ...readBorders(pPr),
  };
  const tabStops = readTabStopDirectives(pPr);
  if (tabStops.length > 0) format.tabStops = tabStops;
  const align = ALIGN_BY_JC[childValue(pPr, "jc") ?? ""];
  if (align) format.align = align;
  if (childByLocalName(pPr, "bidi")) {
    format.direction = isOn(pPr, "bidi") ? "rtl" : "ltr";
  }
  if (isOn(pPr, "pageBreakBefore")) format.pageBreakBefore = true;
  const background = shadingOf(pPr);
  if (background) format.background = background;
  return format;
}

export function readRunFormat(
  rPr: Element | null,
  themeFonts: ThemeFonts = NO_THEME_FONTS
): RunFormat | null {
  if (!rPr) return null;
  const format: RunFormat = {};
  if (isOn(rPr, "b")) format.bold = true;
  if (isOn(rPr, "i")) format.italic = true;
  if (isOn(rPr, "strike")) format.strike = true;
  if (isOn(rPr, "smallCaps")) format.smallCaps = true;

  const underline = childValue(rPr, "u");
  if (isUnderlineKind(underline)) format.underline = underline;

  const fontSizePt = halfPointsToPt(childValue(rPr, "sz"));
  if (fontSizePt !== null) format.fontSizePt = fontSizePt;

  const rFonts = childByLocalName(rPr, "rFonts");
  const fontFamily = rFonts ? fontFamilyOf(rFonts, themeFonts) : null;
  if (fontFamily) format.fontFamily = fontFamily;

  const color = toHexColor(childValue(rPr, "color"));
  if (color) format.color = color;

  const highlight = childValue(rPr, "highlight");
  if (isHighlightName(highlight)) format.highlight = highlight;

  const background = shadingOf(rPr);
  if (background) format.background = background;

  const verticalAlign =
    VERTICAL_ALIGN_BY_VAL[childValue(rPr, "vertAlign") ?? ""];
  if (verticalAlign) format.verticalAlign = verticalAlign;

  const lang = langOf(rPr);
  if (lang) format.lang = lang;

  return format;
}

export const NO_DOCUMENT_DEFAULTS: DocumentDefaults = {
  fontSizePt: null,
  fontFamily: null,
  lineSpacing: null,
};

function defaultProperties(
  styles: Document,
  slot: "rPrDefault" | "pPrDefault",
  name: "rPr" | "pPr"
): Element | null {
  const docDefaults = childByLocalName(styles.documentElement, "docDefaults");
  const wrapper = docDefaults ? childByLocalName(docDefaults, slot) : null;
  return wrapper ? childByLocalName(wrapper, name) : null;
}

/** Reads the paragraph-property layer at the base of the OOXML hierarchy. */
export function readDefaultParagraphFormat(
  styles: Document
): ParagraphFormatLayer {
  return (
    readParagraphFormat(defaultProperties(styles, "pPrDefault", "pPr")) ?? {}
  );
}

/**
 * Reads the document default font size, font, and line spacing from docDefaults in styles.xml.
 * What is absent is left as null, and the on-screen fallback is decided by the display layer.
 */
export function readDocumentDefaults(
  styles: Document,
  themeFonts: ThemeFonts = NO_THEME_FONTS
): DocumentDefaults {
  const rPr = defaultProperties(styles, "rPrDefault", "rPr");
  const pPr = defaultProperties(styles, "pPrDefault", "pPr");
  const rFonts = rPr ? childByLocalName(rPr, "rFonts") : null;
  return {
    fontSizePt: rPr ? halfPointsToPt(childValue(rPr, "sz")) : null,
    fontFamily: rFonts ? fontFamilyOf(rFonts, themeFonts) : null,
    lineSpacing: pPr ? readLineSpacing(pPr) : null,
  };
}
