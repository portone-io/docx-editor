import type { RunFormat } from "../../model/format";
import {
  EMPTY_NUMBERING,
  type Numbering,
  type NumberingOptions,
  parseNumbering,
} from "../../numbering/parseNumbering";
import { type CompatSettings, NO_COMPAT } from "../documentSettings";
import { NO_THEME_FONTS, type ThemeFonts } from "../theme";
import {
  readDefaultParagraphFormat,
  readRunDefaults,
  readRunFormat,
} from "./direct";
import {
  defaultParagraphStyleIdOf,
  defaultTableStyleIdOf,
  NO_STYLES,
  numberingStyleLinks,
  readStyles,
  type StyleTable,
} from "./styles";
import type { ParagraphFormatLayer } from "./tabStops";

/**
 * Everything a document lays down that the §17.7.2 hierarchy is resolved against.
 *
 * Built once when the document is opened and read by every caller that derives a paragraph's or
 * a run's display values, so that a value is the same whichever path asked for it.
 */
export interface FormattingContext {
  styles: StyleTable;
  /** The paragraph style a paragraph naming none wears (`w:default="1"`). Null when the document marks none */
  defaultParagraphStyleId: string | null;
  defaultTableStyleId: string | null;
  /** `pPrDefault`, the layer at the base of the hierarchy */
  paragraphDefaults: ParagraphFormatLayer;
  /**
   * `rPrDefault` read in full. It decides what a removed run property falls back to. The default
   * font and size are supplied through CSS variables; other properties enter derived run values.
   */
  runDefaults: RunFormat;
  numbering: Numbering;
  themeFonts: ThemeFonts;
  compat: CompatSettings;
}

export const NO_FORMATTING: FormattingContext = {
  styles: NO_STYLES,
  defaultParagraphStyleId: null,
  defaultTableStyleId: null,
  paragraphDefaults: {},
  runDefaults: {},
  numbering: EMPTY_NUMBERING,
  themeFonts: NO_THEME_FONTS,
  compat: NO_COMPAT,
};

/**
 * How a document's numbering part is read against the rest of what the document lays down.
 *
 * The numbering folder cannot reach styles.xml, so the links a definition follows are handed to it
 * from here, and the two callers that read a numbering part - opening a document and asking an
 * open one for its lists - resolve the same links.
 */
export function numberingOptionsFor(
  styles: StyleTable,
  themeFonts: ThemeFonts
): NumberingOptions {
  return {
    links: numberingStyleLinks(styles),
    readRun: (rPr) => readRunFormat(rPr, themeFonts),
  };
}

/** The context an opened document resolves against. A document without a styles part lays down nothing */
export function formattingContextOf(
  styles: Document | null,
  numberingXml: string | null,
  themeFonts: ThemeFonts,
  compat: CompatSettings = NO_COMPAT
): FormattingContext {
  const table = styles === null ? NO_STYLES : readStyles(styles, themeFonts);
  const numbering = parseNumbering(
    numberingXml,
    numberingOptionsFor(table, themeFonts)
  );
  if (styles === null) {
    return { ...NO_FORMATTING, numbering, themeFonts, compat };
  }
  return {
    styles: table,
    defaultParagraphStyleId: defaultParagraphStyleIdOf(styles),
    defaultTableStyleId: defaultTableStyleIdOf(styles),
    paragraphDefaults: readDefaultParagraphFormat(styles),
    runDefaults: readRunDefaults(styles, themeFonts),
    numbering,
    themeFonts,
    compat,
  };
}
