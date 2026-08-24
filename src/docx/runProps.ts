/**
 * Changes the character formatting XML (`<w:rPr>`) one item at a time.
 *
 * Turning bold on and off must not lose the other formatting that run wrote down (fonts, shading,
 * even things we do not know about). So instead of building a new fragment we swap out just the one
 * child. It is the same principle `tableFormatting` applies to `w:tcPr`.
 *
 * The display values produced here are used on screen only. What goes back into the document is
 * always the rPr string.
 */

import { type RunFormat, toRunFormat } from "../model/format";
import { normalizeHex } from "../ooxml/units";
import { escapeXml, localPart } from "../ooxml/xml";
import { isEastAsianFontName } from "../styles/fontStack";
import { fontNamesOf, readRunFormat } from "./formatting";
import {
  type Props,
  parseProps,
  parsePropsXml,
  propsChild,
  RUN_PR_ORDER,
  renderProps,
  setPropsChild,
} from "./propsXml";
import { THEME_ATTRS } from "./theme";

/** The character formatting that is toggled on and off */
export type RunToggle = "bold" | "italic" | "underline" | "strike";

/** A job that changes one piece of character formatting. A null value means withdrawing the setting */
export type RunEdit =
  | { kind: "toggle"; toggle: RunToggle; on: boolean }
  | { kind: "fontSize"; pt: number | null }
  | { kind: "fontFamily"; name: string | null }
  | { kind: "color"; hex: string | null }
  | { kind: "background"; hex: string | null };

/** The formatting a run holds. The original XML and the display values derived from that XML form a pair */
export interface RunProps {
  /** The whole `<w:rPr>...</w:rPr>` XML. null for a run with no formatting */
  rPr: string | null;
  format: RunFormat | null;
}

/**
 * The elements used when turning something on.
 * For bold and italic we write the Latin and complex-script pair together, following the convention
 * of Korean documents.
 */
const TOGGLE_CHILDREN: Record<RunToggle, readonly string[]> = {
  bold: ["b", "bCs"],
  italic: ["i", "iCs"],
  underline: ["u"],
  strike: ["strike"],
};

/** The value written to mean on. Only underline also records a kind */
const TOGGLE_ON_VALUE: Record<RunToggle, string | null> = {
  bold: null,
  italic: null,
  underline: "single",
  strike: null,
};

/** The value written when pinning down an off state */
const TOGGLE_OFF_VALUE: Record<RunToggle, string> = {
  bold: "0",
  italic: "0",
  underline: "none",
  strike: "0",
};

/** The limit on the font size Word records in half-points (819pt) */
const MAX_FONT_SIZE_PT = 819;

function elementXml(name: string, value: string | null): string {
  return value === null ? `<w:${name}/>` : `<w:${name} w:val="${value}"/>`;
}

/**
 * The shading that records a text background. `clear` means paint with the `fill` color alone, with no pattern.
 * `fill="auto"` means "no background", so it acts as an off that overrides inheritance.
 */
function shadingXml(fill: string): string {
  return `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
}

/** Gathers `#2e74b5` and `2E74B5` alike into the shape the document uses. null if it is not a color */
/** Whether the size can be recorded in half-points. Anything not on a 0.5pt step cannot be written into the document */
function halfPoints(pt: number): number | null {
  if (!Number.isFinite(pt) || pt <= 0 || pt > MAX_FONT_SIZE_PT) return null;
  const half = pt * 2;
  return Number.isInteger(half) ? half : null;
}

/**
 * The slots that record font names.
 *
 * A single run records the Latin, the East Asian, and the complex-script font apart from one
 * another, and which of them a chosen font is written into depends on the font itself.
 * An East Asian name goes into all three: it is one name meant for the whole run, which is
 * how the Korean contract fixtures write it and what Word writes when the font picked is a
 * CJK one.
 * A Latin name goes into the Latin slots alone and leaves the East Asian slot standing.
 * A Japanese or Chinese document names a Latin font beside its own on purpose - the Latin one
 * for the letters and digits, its own for the rest - so writing the Latin name into the East
 * Asian slot as well would draw its Japanese text in a font that has no such glyphs at all.
 */
const LATIN_SLOTS: readonly string[] = ["ascii", "hAnsi"];
const EAST_ASIAN_SLOTS: readonly string[] = [...LATIN_SLOTS, "eastAsia"];

/**
 * The slots one name is written into.
 * The complex-script slot follows along wherever the run already carries one; a run carrying
 * none is not given one.
 */
function fontSlots(name: string, hasComplexScript: boolean): readonly string[] {
  const slots = isEastAsianFontName(name) ? EAST_ASIAN_SLOTS : LATIN_SLOTS;
  return hasComplexScript ? [...slots, "cs"] : slots;
}

/** Whether the font name can be written as is into both the document and the screen. null otherwise */
function fontName(value: string): string | null {
  const name = value.trim();
  // The display value is a CSS name list, so a name holding a quote or a semicolon breaks the declaration
  return name.length > 0 && !/["';<>&]/.test(name) ? name : null;
}

/** The attributes of the rFonts this run wrote down. null if its shape cannot be made out */
function rFontsAttrs(xml: string | null): [string, string][] | null {
  if (xml === null) return [];
  const el = parsePropsXml(xml);
  if (!el) return null;
  return Array.from(el.attributes, (attr): [string, string] => [
    attr.name,
    attr.value,
  ]);
}

/**
 * A new rFonts with the font name written in.
 *
 * The name goes into the slots `fontSlots` names, and a theme reference in one of those slots
 * is cleared away, since a theme beats a name. The slots the name does not reach keep both
 * their name and their theme reference.
 * The attributes we do not decide (`w:hint` and so on) keep their original values.
 */
function rFontsXml(current: string | null, name: string): string | null {
  const attrs = rFontsAttrs(current);
  if (!attrs) return null;
  const slots = fontSlots(
    name,
    attrs.some(([attr]) => localPart(attr) === "cs")
  );
  const dropped = slots.flatMap((slot) => [slot, ...(THEME_ATTRS[slot] ?? [])]);
  const kept = attrs.filter(([attr]) => !dropped.includes(localPart(attr)));
  const text = [
    ...slots.map((slot): [string, string] => [`w:${slot}`, name]),
    ...kept,
  ]
    .map(([attr, value]) => `${attr}="${escapeXml(value)}"`)
    .join(" ");
  return `<w:rFonts ${text}/>`;
}

function hasChild(xml: string | null, name: string): boolean {
  if (xml === null) return false;
  const props = parseProps(xml);
  return props !== null && propsChild(props.children, name) !== undefined;
}

/** Turning formatting on for a run with no formatting creates a minimal rPr from scratch */
function propsOf(rPr: string | null): Props | null {
  if (rPr === null) return { tag: "w:rPr", attrs: null, children: [] };
  return parseProps(rPr);
}

/** What text one child is to be changed to. A null xml removes that child */
type ChildEdit = readonly [name: string, xml: string | null];

/**
 * How formatting is turned off.
 *
 * Where the setting could arrive already on from a style, removing the element would let the style
 * win again, so we pin the off state down instead. With no inheritance (most of the fixtures),
 * removing the element is what comes closest to the original.
 */
function offEdit(name: string, value: string, inherited: boolean): ChildEdit {
  return [name, inherited ? elementXml(name, value) : null];
}

/** Which children of the rPr one job changes and how. null for a value whose meaning cannot be made out */
function childEdits(
  edit: RunEdit,
  inherited: boolean,
  rFonts: string | null
): ChildEdit[] | null {
  switch (edit.kind) {
    case "toggle": {
      const names = TOGGLE_CHILDREN[edit.toggle];
      if (edit.on) {
        const value = TOGGLE_ON_VALUE[edit.toggle];
        return names.map((name) => [name, elementXml(name, value)]);
      }
      const off = TOGGLE_OFF_VALUE[edit.toggle];
      return names.map((name) => offEdit(name, off, inherited));
    }
    case "fontSize": {
      // There is no off for size. Withdrawing the setting inherits the style and the document default again
      if (edit.pt === null) {
        return [
          ["sz", null],
          ["szCs", null],
        ];
      }
      const half = halfPoints(edit.pt);
      if (half === null) return null;
      return [
        ["sz", elementXml("sz", `${half}`)],
        ["szCs", elementXml("szCs", `${half}`)],
      ];
    }
    case "fontFamily": {
      // There is no off for the font either. Withdrawing the setting removes the whole rFonts so the document default font applies again
      if (edit.name === null) return [["rFonts", null]];
      const name = fontName(edit.name);
      if (name === null) return null;
      const xml = rFontsXml(rFonts, name);
      return xml === null ? null : [["rFonts", xml]];
    }
    case "color": {
      // auto means "the text color as the document decides", so it acts as an off that overrides inheritance
      if (edit.hex === null) return [offEdit("color", "auto", inherited)];
      const hex = normalizeHex(edit.hex);
      return hex === null ? null : [["color", elementXml("color", hex)]];
    }
    case "background": {
      // Word paints the highlight on top of the shading. Left in place, it would hide the new background color underneath it
      const highlight = offEdit("highlight", "none", inherited);
      if (edit.hex === null) {
        return [["shd", inherited ? shadingXml("auto") : null], highlight];
      }
      const hex = normalizeHex(edit.hex);
      return hex === null ? null : [["shd", shadingXml(hex)], highlight];
    }
  }
}

/** Derives the display values again from the rPr we operated on. The same rPr always yields the same display values */
export function readRunProps(rPr: string | null): RunFormat | null {
  return rPr === null ? null : readRunFormat(parsePropsXml(rPr));
}

/**
 * The part of the display values that came from the paragraph style rather than from the rPr.
 * Our surgery only touches the rPr, so the values the style gave are left as they are.
 */
function inheritedFormat(current: RunProps): Record<string, unknown> {
  const direct = new Set(Object.keys(readRunProps(current.rPr) ?? {}));
  const inherited: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(current.format ?? {})) {
    if (!direct.has(key)) inherited[key] = value;
  }
  return inherited;
}

function nextProps(rPr: string | null, previous: RunProps): RunProps {
  const inherited = inheritedFormat(previous);
  const format = toRunFormat({ ...inherited, ...(readRunProps(rPr) ?? {}) });
  // A run left with no formatting at all goes back to its initial state, with neither an rPr nor display values
  if (rPr === null && (format === null || Object.keys(format).length === 0)) {
    return { rPr: null, format: null };
  }
  return { rPr, format };
}

/**
 * The rPr and display values after changing one piece of character formatting.
 *
 * `pPr` is used only to see whether the paragraph points at a style.
 * null for an rPr whose shape could not be made out, or for a value that cannot be written into
 * the document, in which case the original is left untouched.
 */
export function editRunProps(
  current: RunProps,
  pPr: string | null,
  edit: RunEdit
): RunProps | null {
  const props = propsOf(current.rPr);
  if (!props) return null;

  const inherited =
    propsChild(props.children, "rStyle") !== undefined ||
    hasChild(pPr, "pStyle");
  const rFonts = propsChild(props.children, "rFonts")?.xml ?? null;
  const edits = childEdits(edit, inherited, rFonts);
  if (!edits) return null;

  const children = edits.reduce(
    (kept, [name, xml]) => setPropsChild(kept, name, xml, RUN_PR_ORDER),
    props.children
  );
  const rPr = renderProps({ ...props, children });
  return nextProps(rPr === "" ? null : rPr, current);
}

/** Whether one of the toggled formats is on in these display values */
export function isRunToggleOn(
  format: RunFormat | null,
  toggle: RunToggle
): boolean {
  if (!format) return false;
  if (toggle === "bold") return format.bold === true;
  if (toggle === "italic") return format.italic === true;
  if (toggle === "strike") return format.strike === true;
  // Underline holds a kind rather than an on/off state
  return format.underline !== undefined;
}

/** Whether this text is already in the state the job wants. If it already is, we leave it untouched */
export function matchesRunEdit(
  format: RunFormat | null,
  edit: RunEdit
): boolean {
  switch (edit.kind) {
    case "toggle":
      return isRunToggleOn(format, edit.toggle) === edit.on;
    case "fontSize":
      return (format?.fontSizePt ?? null) === edit.pt;
    case "fontFamily": {
      const names = fontNamesOf(format?.fontFamily);
      if (edit.name === null) return names.length === 0;
      // If the slots are using different names, it is not yet in the state we want
      return names.length === 1 && names[0] === fontName(edit.name);
    }
    case "color": {
      const current = format?.color ?? null;
      if (current === null || edit.hex === null) return current === edit.hex;
      return normalizeHex(current) === normalizeHex(edit.hex);
    }
    case "background": {
      // If a highlight is still there, it is not yet in the state we want. It has to move over to shading
      if (format?.highlight !== undefined) return false;
      const fill = format?.background ? normalizeHex(format.background) : null;
      return fill === (edit.hex === null ? null : normalizeHex(edit.hex));
    }
  }
}
