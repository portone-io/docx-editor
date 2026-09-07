import { NO_FILL, type ParagraphAlign } from "../model/format";
import {
  EIGHTHS_PER_PT,
  HALF_POINTS_PER_PT,
  type HexColor,
  ST_EighthPointMeasure,
  ST_HexColor,
  ST_OnOff,
  TWIPS_PER_PT,
} from "./simpleTypes";
import { childByLocalName, W_NS } from "./xml";

/** The alignment `w:jc` writes down, both for a paragraph and for a table */
export const ALIGN_BY_JC: Record<string, ParagraphAlign> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "justify",
  distribute: "justify",
};

/** A color as the document writes it, six hex digits and no `#`. Null for anything else */
export function normalizeHex(value: string): string | null {
  const hex = value.startsWith("#") ? value.slice(1) : value;
  return /^[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : null;
}

export function wAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(W_NS, name) ?? el.getAttribute(name);
}

/**
 * Whether a boolean property element states on (§17.17.4).
 *
 * An element carrying no `w:val` states on, and one that is not there at all states nothing, which
 * a caller asking "is this on" reads as off. For malformed values, this project uses the same
 * fallback as an absent attribute. That recovery policy is not a schema or Word guarantee.
 */
export function isOnElement(el: Element | null): boolean {
  if (!el) return false;
  return ST_OnOff.parse(wAttr(el, "val")) ?? true;
}

/** `<w:b/>` means on, `<w:b w:val="0"/>` means off */
export function isOn(parent: Element, name: string): boolean {
  return isOnElement(childByLocalName(parent, name));
}

export function childValue(parent: Element, name: string): string | null {
  const el = childByLocalName(parent, name);
  return el ? wAttr(el, "val") : null;
}

/** Cuts off the messy digits after the decimal point so the same input always yields the same value */
export function round(value: number): number {
  return Math.round(value * 100) / 100;
}

/** twip (1/20 of a point) */
export function twipsToPt(twips: number | null): number | null {
  return twips === null ? null : round(twips / TWIPS_PER_PT);
}

/** half-point (w:sz 20 = 10pt) */
export function halfPointsToPt(half: number | null): number | null {
  return half === null ? null : round(half / HALF_POINTS_PER_PT);
}

/** The color as CSS spells it. `auto` is not a color to paint with, so it reads as none */
export function toHexColor(value: string | null): string | null {
  const color = ST_HexColor.parse(value);
  return color?.kind === "rgb" ? `#${color.hex}` : null;
}

/**
 * Every line style `ST_Border` names, rounded to the four kinds CSS gives us: a line drawn once is
 * solid, a line drawn as two or more strokes is double, and a broken line is dotted or dashed.
 *
 * The decorative art borders (`apples` and the rest, which only a page border may carry) are left
 * out on purpose, so a side wearing one reads as "we do not know this kind".
 */
const BORDER_STYLE_BY_VAL: Record<string, string> = {
  single: "solid",
  thick: "solid",
  double: "double",
  dotted: "dotted",
  dashed: "dashed",
  dotDash: "dashed",
  wave: "solid",
  outset: "solid",
  inset: "solid",
  threeDEmboss: "solid",
  threeDEngrave: "solid",
  triple: "double",
  doubleWave: "double",
  thinThickSmallGap: "double",
  thickThinSmallGap: "double",
  thinThickThinSmallGap: "double",
  thinThickMediumGap: "double",
  thickThinMediumGap: "double",
  thinThickThinMediumGap: "double",
  thinThickLargeGap: "double",
  thickThinLargeGap: "double",
  thinThickThinLargeGap: "double",
  dashSmallGap: "dashed",
  dotDotDash: "dashed",
  dashDotStroked: "dashed",
};

/**
 * The ST_Border value each of those four CSS styles stands for.
 *
 * It is the way back from a display value to markup, which is what a caller holding a line as CSS
 * rather than as the element it was read from has to make before it can write that line down. The
 * two tables sit together so that neither can name a style the other does not.
 */
const BORDER_VAL_BY_STYLE: Readonly<Record<string, string>> = {
  solid: "single",
  double: "double",
  dashed: "dashed",
  dotted: "dotted",
};

/** What Word draws for a side that names a line but no thickness */
const FALLBACK_BORDER_EIGHTHS = 4;

/** The color a line whose own is `auto` is drawn in */
const DEFAULT_BORDER_CSS_COLOR = "#000000";

/** One border side as `CT_Border` records it (§17.3.4) */
export interface BorderLine {
  /** ST_Border (§17.18.2). `none` states that this side draws nothing */
  val: string;
  /** The thickness, in eighths of a point */
  eighths: number;
  color: HexColor;
}

const NO_LINE: BorderLine = {
  val: "none",
  eighths: 0,
  color: { kind: "auto" },
};

/**
 * The line one border side draws.
 *
 * A side the document has pinned down as "draw no line here" draws `none`.
 * A kind we do not know, or a side that is not there at all, is null, so the caller can pick a default.
 * A side that draws a line but writes down no thickness gets the default one: reading it as "no
 * line" would take away a line the style laid down.
 */
function borderLineOf(el: Element | null): BorderLine | null {
  if (!el) return null;
  const val = wAttr(el, "val") ?? "";
  if (val === "nil" || val === "none") return NO_LINE;
  if (!BORDER_STYLE_BY_VAL[val]) return null;
  const written = ST_EighthPointMeasure.parse(wAttr(el, "sz"));
  return {
    val,
    eighths:
      written === null || written === 0 ? FALLBACK_BORDER_EIGHTHS : written,
    color: ST_HexColor.parse(wAttr(el, "color")) ?? { kind: "auto" },
  };
}

/** The CSS declaration one line is drawn with, and null for a line there is none of */
export function borderLineCss(line: BorderLine | null): string | null {
  if (!line) return null;
  if (line.val === "none") return "none";
  const style = BORDER_STYLE_BY_VAL[line.val];
  if (!style) return null;
  const color =
    line.color.kind === "rgb" ? `#${line.color.hex}` : DEFAULT_BORDER_CSS_COLOR;
  return `${round(line.eighths / EIGHTHS_PER_PT)}pt ${style} ${color}`;
}

const BORDER_CSS =
  /^(\d+(?:\.\d+)?)pt (solid|double|dashed|dotted) (#[0-9a-f]{6})$/i;

/** The line a CSS declaration this module wrote stands for, so no caller reads one back its own way */
export function borderLineOfCss(css: string | null): BorderLine | null {
  if (css === null) return null;
  if (css === "none") return NO_LINE;
  const matched = BORDER_CSS.exec(css);
  if (!matched) return null;
  const eighths = Math.round(Number(matched[1]) * EIGHTHS_PER_PT);
  const color = ST_HexColor.parse(matched[3].slice(1).toUpperCase());
  if (eighths <= 0 || color === null) return null;
  return { val: BORDER_VAL_BY_STYLE[matched[2].toLowerCase()], eighths, color };
}

/** Moves one border side into a CSS value */
export function borderCss(el: Element | null): string | null {
  return borderLineCss(borderLineOf(el));
}

/** One side inside `w:tblBorders` or `w:tcBorders` */
export function borderSide(
  borders: Element | null,
  side: string
): string | null {
  return borders ? borderCss(childByLocalName(borders, side)) : null;
}

/**
 * The background color from `w:shd`. For a pattern fill we only look at the color.
 *
 * A shading with no color to fill with (`w:fill="auto"`, or none at all) says to paint nothing, so
 * it reads back as `NO_FILL`. That is what takes a fill a style laid down back off, the way Word does.
 */
export function shadingOf(parent: Element): string | null {
  const shd = childByLocalName(parent, "shd");
  if (!shd) return null;
  return toHexColor(wAttr(shd, "fill")) ?? NO_FILL;
}
