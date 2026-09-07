/**
 * Which attribute beats which when an element writes down both.
 *
 * OOXML records the same setting twice over in several places: a font slot beside the theme
 * reference that overrules it, a color beside the theme color, an indent in twips beside the same
 * indent counted in characters. Writing one of a pair while leaving the other standing means
 * writing a value Word will not show, so the one that would win is dropped as the other is written.
 *
 * The table is keyed by the local names of the owning element and the attribute together, because
 * neither is enough on its own: `color` is an attribute of `w:shd` but an element of its own whose
 * value sits in `w:val`, and `left` is an attribute of `w:ind` but an element of `w:tcMar`.
 */

import { withAttr, withoutAttrs, type XmlAttr } from "./element";

/** The theme reference that beats a color, wherever an element records one (§17.3.5) */
const THEME_COLOR: readonly string[] = [
  "themeColor",
  "themeTint",
  "themeShade",
];

const NO_OVERRIDES: readonly string[] = [];

/**
 * The attributes that override the named one when both are there, so writing the named one drops
 * them (§17.3.1.12 ind, §17.3.2.26 rFonts, §17.3.2.6 color, §17.3.5 shd, §17.3.1.33 spacing).
 *
 * The border sides are the children of `w:tcBorders`, each of them a CT_Border that records a
 * color of its own.
 */
export const OVERRIDING_ATTRS: Readonly<
  Record<`${string}/${string}`, readonly string[]>
> = {
  "rFonts/ascii": ["asciiTheme"],
  "rFonts/hAnsi": ["hAnsiTheme"],
  "rFonts/eastAsia": ["eastAsiaTheme"],
  // This is the slot whose capitalization varies from document to document
  "rFonts/cs": ["cstheme", "csTheme"],
  "color/val": THEME_COLOR,
  "shd/color": THEME_COLOR,
  "shd/fill": ["themeFill", "themeFillTint", "themeFillShade"],
  "ind/left": ["leftChars"],
  "ind/start": ["startChars"],
  "ind/right": ["rightChars"],
  "ind/end": ["endChars"],
  "ind/hanging": ["hangingChars"],
  "ind/firstLine": ["firstLineChars"],
  "spacing/before": ["beforeLines"],
  "spacing/after": ["afterLines"],
  "top/color": THEME_COLOR,
  "start/color": THEME_COLOR,
  "left/color": THEME_COLOR,
  "bottom/color": THEME_COLOR,
  "end/color": THEME_COLOR,
  "right/color": THEME_COLOR,
  "insideH/color": THEME_COLOR,
  "insideV/color": THEME_COLOR,
  "tl2br/color": THEME_COLOR,
  "tr2bl/color": THEME_COLOR,
};

/** What writing `element`'s `name` drops along with it. Empty for a pair nothing overrides */
export function overridingAttrs(
  element: string,
  name: string
): readonly string[] {
  return OVERRIDING_ATTRS[`${element}/${name}`] ?? NO_OVERRIDES;
}

/**
 * `withAttr` with the attributes the new value would be fighting with dropped as it is written.
 *
 * This is how an attribute is meant to be written, so that a writer added later cannot forget the
 * drop. A caller that means to leave an overriding attribute standing says so by reaching for
 * `withAttr` instead.
 */
export function setAttr(
  attrs: readonly XmlAttr[],
  element: string,
  name: string,
  value: string | null
): XmlAttr[] {
  return withAttr(
    withoutAttrs(attrs, overridingAttrs(element, name)),
    name,
    value
  );
}
