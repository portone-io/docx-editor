import { NO_FILL, type ParagraphAlign } from "../model/format";
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

/** `<w:b/>` means on, `<w:b w:val="0"/>` means off */
export function isOn(parent: Element, name: string): boolean {
  const el = childByLocalName(parent, name);
  if (!el) return false;
  const value = wAttr(el, "val");
  return value === null || value === "1" || value === "true" || value === "on";
}

export function childValue(parent: Element, name: string): string | null {
  const el = childByLocalName(parent, name);
  return el ? wAttr(el, "val") : null;
}

export function toNumber(value: string | null): number | null {
  if (value === null) return null;
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : null;
}

/** Cuts off the messy digits after the decimal point so the same input always yields the same value */
export function round(value: number): number {
  return Math.round(value * 100) / 100;
}

/** twip (1/20 of a point) */
export function twipsToPt(value: string | null): number | null {
  const twips = toNumber(value);
  return twips === null ? null : round(twips / 20);
}

/** half-point (w:sz 20 = 10pt) */
export function halfPointsToPt(value: string | null): number | null {
  const half = toNumber(value);
  return half === null ? null : round(half / 2);
}

/** 1/8 of a point */
export function eighthsToPt(value: string | null): number | null {
  const eighths = toNumber(value);
  return eighths === null ? null : round(eighths / 8);
}

export function toHexColor(value: string | null): string | null {
  if (value === null || value === "auto") return null;
  return /^[0-9a-fA-F]{6}$/.test(value) ? `#${value}` : null;
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

/** What Word draws for a side that names a line but no thickness */
const FALLBACK_BORDER_PT = 0.5;

/**
 * Moves one border side into a CSS value.
 *
 * A side the document has pinned down as "draw no line here" becomes `none`.
 * A kind we do not know, or a side that is not there at all, becomes null, so the caller can pick a default.
 * A side that draws a line but writes down no thickness gets the default one: reading it as "no
 * line" would take away a line the style laid down.
 */
export function borderCss(el: Element | null): string | null {
  if (!el) return null;
  const val = wAttr(el, "val") ?? "";
  if (val === "nil" || val === "none") return "none";
  const style = BORDER_STYLE_BY_VAL[val];
  if (!style) return null;
  const written = eighthsToPt(wAttr(el, "sz"));
  const widthPt =
    written === null || written === 0 ? FALLBACK_BORDER_PT : written;
  const color = toHexColor(wAttr(el, "color")) ?? "#000000";
  return `${widthPt}pt ${style} ${color}`;
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
