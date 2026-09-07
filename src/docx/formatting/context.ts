import type { RunFormat } from "../../model/format";
import {
  EMPTY_NUMBERING,
  type Numbering,
} from "../../numbering/parseNumbering";
import { type CompatSettings, NO_COMPAT } from "../documentSettings";
import { NO_THEME_FONTS, type ThemeFonts } from "../theme";
import { readDefaultParagraphFormat, readRunDefaults } from "./direct";
import {
  defaultParagraphStyleIdOf,
  defaultTableStyleIdOf,
  NO_STYLES,
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
   * `rPrDefault` read in full. It decides what a removed run property falls back to; the screen
   * draws it through the CSS variables `DocumentDefaults` fills, not through a run's values
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

/** The context an opened document resolves against. A document without a styles part lays down nothing */
export function formattingContextOf(
  styles: Document | null,
  numbering: Numbering,
  themeFonts: ThemeFonts,
  compat: CompatSettings = NO_COMPAT
): FormattingContext {
  if (styles === null) {
    return { ...NO_FORMATTING, numbering, themeFonts, compat };
  }
  return {
    styles: readStyles(styles, themeFonts),
    defaultParagraphStyleId: defaultParagraphStyleIdOf(styles),
    defaultTableStyleId: defaultTableStyleIdOf(styles),
    paragraphDefaults: readDefaultParagraphFormat(styles),
    runDefaults: readRunDefaults(styles, themeFonts),
    numbering,
    themeFonts,
    compat,
  };
}
