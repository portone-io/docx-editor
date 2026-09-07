/**
 * Reads direct formatting XML into display values.
 *
 * The values produced here are used for display only.
 * What goes back into the document is always the original XML string, so no matter what is read
 * here, export is undisturbed.
 */

import {
  type DocumentDefaults,
  type LineSpacing,
  type ParagraphFormat,
  RUN_FORMAT_KEYS,
  type RunFormat,
} from "../../model/format";
import {
  ST_DecimalNumber,
  ST_SignedTwipsMeasure,
  ST_TwipsMeasure,
  TWIPS_PER_PT,
} from "../../ooxml/simpleTypes";
import { readTabStopDirectives } from "../../ooxml/tabStops";
import {
  ALIGN_BY_JC,
  borderSide,
  childValue,
  isOn,
  round,
  shadingOf,
  twipsToPt,
  wAttr,
} from "../../ooxml/units";
import { childByLocalName } from "../../ooxml/xml";
import { NO_THEME_FONTS, type ThemeFonts } from "../theme";
import { RUN_PROPERTIES } from "./runProperties";
import type { ParagraphFormatLayer } from "./tabStops";

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
  const start = twipsToPt(
    ST_SignedTwipsMeasure.parse(wAttr(ind, "start") ?? wAttr(ind, "left"))
  );
  if (start !== null) format.indentStartPt = start;
  const end = twipsToPt(
    ST_SignedTwipsMeasure.parse(wAttr(ind, "end") ?? wAttr(ind, "right"))
  );
  if (end !== null) format.indentEndPt = end;
  // A hanging indent overrides the first-line indent (an OOXML rule)
  const hanging = twipsToPt(ST_TwipsMeasure.parse(wAttr(ind, "hanging")));
  const firstLine = twipsToPt(ST_TwipsMeasure.parse(wAttr(ind, "firstLine")));
  if (hanging !== null) format.textIndentPt = -hanging;
  else if (firstLine !== null) format.textIndentPt = firstLine;
  return format;
}

/** `w:spacing/@line` counts an automatic line height in 240ths of a line (§17.3.1.33) */
export const LINE_UNITS_PER_LINE = 240;

/**
 * The line spacing from `w:spacing`.
 * auto means a multiple in 240ths, and the rest are heights pinned down in twips.
 */
export function readLineSpacing(pPr: Element): LineSpacing | null {
  const spacing = childByLocalName(pPr, "spacing");
  if (!spacing) return null;
  const line = ST_SignedTwipsMeasure.parse(wAttr(spacing, "line"));
  if (line === null) return null;
  const rule = wAttr(spacing, "lineRule");
  const pt = round(line / TWIPS_PER_PT);
  if (rule === "exact") return { rule: "exact", pt };
  if (rule === "atLeast") return { rule: "atLeast", pt };
  return { rule: "auto", lines: round(line / LINE_UNITS_PER_LINE) };
}

function readSpacing(pPr: Element): Partial<ParagraphFormat> {
  const spacing = childByLocalName(pPr, "spacing");
  if (!spacing) return {};
  const format: Partial<ParagraphFormat> = {};
  const before = twipsToPt(ST_TwipsMeasure.parse(wAttr(spacing, "before")));
  if (before !== null) format.spaceBeforePt = before;
  const after = twipsToPt(ST_TwipsMeasure.parse(wAttr(spacing, "after")));
  if (after !== null) format.spaceAfterPt = after;
  const lineSpacing = readLineSpacing(pPr);
  if (lineSpacing) format.lineSpacing = lineSpacing;
  return format;
}

/** The list slot `w:numPr` points at. A numId of 0 means "not a list" */
function readNumbering(pPr: Element): Partial<ParagraphFormatLayer> {
  const numPr = childByLocalName(pPr, "numPr");
  if (!numPr) return {};
  const numId = ST_DecimalNumber.parse(childValue(numPr, "numId"));
  if (numId === 0) return { numbering: null };
  if (numId === null || numId < 0) return {};
  return {
    numbering: {
      numId,
      ilvl: ST_DecimalNumber.parse(childValue(numPr, "ilvl")) ?? 0,
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
  if (isOn(pPr, "keepNext")) format.keepNext = true;
  const background = shadingOf(pPr);
  if (background) format.background = background;
  return format;
}

function readRunProperty<K extends keyof RunFormat>(
  format: RunFormat,
  key: K,
  rPr: Element,
  themeFonts: ThemeFonts
): void {
  const value = RUN_PROPERTIES[key].read(rPr, themeFonts);
  if (value !== undefined) format[key] = value;
}

export function readRunFormat(
  rPr: Element | null,
  themeFonts: ThemeFonts = NO_THEME_FONTS
): RunFormat | null {
  if (!rPr) return null;
  const format: RunFormat = {};
  for (const key of RUN_FORMAT_KEYS) {
    readRunProperty(format, key, rPr, themeFonts);
  }
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

/** Reads the run-property layer at the base of the OOXML hierarchy (`rPrDefault`) in full */
export function readRunDefaults(
  styles: Document,
  themeFonts: ThemeFonts = NO_THEME_FONTS
): RunFormat {
  return (
    readRunFormat(defaultProperties(styles, "rPrDefault", "rPr"), themeFonts) ??
    {}
  );
}

/** The character style a run's formatting points at (`w:rStyle`). null when it points at none */
export function runStyleIdOf(rPr: Element | null): string | null {
  const id = rPr ? childValue(rPr, "rStyle") : null;
  return id !== null && id.length > 0 ? id : null;
}

/**
 * Reads the document default font size, font, and line spacing from docDefaults in styles.xml.
 * What is absent is left as null, and the on-screen fallback is decided by the display layer.
 */
export function readDocumentDefaults(
  styles: Document,
  themeFonts: ThemeFonts = NO_THEME_FONTS
): DocumentDefaults {
  const run = readRunDefaults(styles, themeFonts);
  const pPr = defaultProperties(styles, "pPrDefault", "pPr");
  return {
    fontSizePt: run.fontSizePt ?? null,
    fontFamily: run.fontFamily ?? null,
    lineSpacing: pPr ? readLineSpacing(pPr) : null,
  };
}
