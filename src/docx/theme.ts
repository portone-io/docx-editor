/**
 * The fonts a document's theme names, and the references a run makes to them.
 *
 * Instead of naming a font, a `w:rFonts` may point at the theme
 * (`w:asciiTheme="minorHAnsi"`), which is what Word writes for most of what it produces
 * and what a Japanese or Chinese document almost always carries.
 * Resolving those references is what lets such a run be drawn in the font the reader
 * would see in Word. It is used for display only: the theme attributes themselves are
 * carried back out untouched, so nothing read here reaches the file.
 */

import { childByLocalName } from "../ooxml/xml";

/** The typefaces one font scheme names, one per script */
interface ThemeFontScheme {
  latin: string | null;
  eastAsia: string | null;
  complexScript: string | null;
}

/** The two font schemes a theme defines: one for headings, one for body text */
export interface ThemeFonts {
  major: ThemeFontScheme;
  minor: ThemeFontScheme;
}

const NO_FONT_SCHEME: ThemeFontScheme = {
  latin: null,
  eastAsia: null,
  complexScript: null,
};

/** What applies to a document with no theme part, or one whose theme names no font */
export const NO_THEME_FONTS: ThemeFonts = {
  major: NO_FONT_SCHEME,
  minor: NO_FONT_SCHEME,
};

/**
 * The theme attribute paired with each font slot.
 * A theme reference and a name in the same slot cannot both apply, so the name wins,
 * the way Word reads them.
 */
export const THEME_ATTRS: Record<string, readonly string[]> = {
  ascii: ["asciiTheme"],
  hAnsi: ["hAnsiTheme"],
  eastAsia: ["eastAsiaTheme"],
  // This is the slot whose capitalization varies from document to document
  cs: ["cstheme", "csTheme"],
};

interface ThemeSlot {
  scheme: keyof ThemeFonts;
  script: keyof ThemeFontScheme;
}

/**
 * The scheme and script each value of `ST_Theme` points at.
 * `Ascii` and `HAnsi` both name the Latin typeface; `Bidi` names the complex-script one.
 * The keys are lower case because the value is written by hand often enough that its
 * capitalization cannot be relied on.
 */
const THEME_SLOTS: Record<string, ThemeSlot> = {
  majorascii: { scheme: "major", script: "latin" },
  majorhansi: { scheme: "major", script: "latin" },
  majoreastasia: { scheme: "major", script: "eastAsia" },
  majorbidi: { scheme: "major", script: "complexScript" },
  minorascii: { scheme: "minor", script: "latin" },
  minorhansi: { scheme: "minor", script: "latin" },
  minoreastasia: { scheme: "minor", script: "eastAsia" },
  minorbidi: { scheme: "minor", script: "complexScript" },
};

/** The font one theme reference stands for. null for a reference we cannot resolve */
export function themeFontName(
  fonts: ThemeFonts,
  value: string | null
): string | null {
  if (value === null) return null;
  const slot = THEME_SLOTS[value.trim().toLowerCase()];
  return slot ? fonts[slot.scheme][slot.script] : null;
}

/** A typeface a font scheme names. An empty `typeface` means the theme names none */
function typefaceOf(font: Element | null, name: string): string | null {
  const el = font ? childByLocalName(font, name) : null;
  const typeface = el?.getAttribute("typeface") ?? null;
  return typeface !== null && typeface.length > 0 ? typeface : null;
}

function fontSchemeOf(
  fontScheme: Element | null,
  slot: "majorFont" | "minorFont"
): ThemeFontScheme {
  const font = fontScheme ? childByLocalName(fontScheme, slot) : null;
  if (!font) return NO_FONT_SCHEME;
  return {
    latin: typefaceOf(font, "latin"),
    eastAsia: typefaceOf(font, "ea"),
    complexScript: typefaceOf(font, "cs"),
  };
}

/**
 * Reads the font scheme out of theme1.xml.
 *
 * The per-script fonts (`a:font script="Jpan"`) a theme lists beside those three are the
 * ones Word swaps in when the language of the text changes, which is not something the
 * screen decides, so they are left alone.
 */
export function readThemeFonts(theme: Document): ThemeFonts {
  const elements = childByLocalName(theme.documentElement, "themeElements");
  const fontScheme = elements ? childByLocalName(elements, "fontScheme") : null;
  return {
    major: fontSchemeOf(fontScheme, "majorFont"),
    minor: fontSchemeOf(fontScheme, "minorFont"),
  };
}
