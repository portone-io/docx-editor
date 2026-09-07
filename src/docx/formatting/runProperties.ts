/**
 * One row per run property: what it reads out of an rPr, what it writes into one, and whether
 * text already carries a value.
 *
 * Reading, writing and comparing used to be three hand-kept copies of the same vocabulary, so a
 * property could be read under one name and written under another. Here each is one row, and the
 * readers and writers below walk the table, so anything else that has to speak for a run
 * property - a tracked change's `w:rPrChange` snapshot, a style's `w:rPr` - asks the same row.
 *
 * The values read are for display only. What goes back into the document is the rPr string, and
 * a write is a list of children swapped into it (`ooxml/props`), never a tree serialized again.
 */

import {
  isHighlightName,
  isUnderlineKind,
  NO_FILL,
  RUN_FORMAT_KEYS,
  type RunFormat,
  type VerticalAlign,
} from "../../model/format";
import { elementXml, wAttrValue, type XmlAttr } from "../../ooxml/element";
import { wName } from "../../ooxml/names";
import { overridingAttrs, setAttr } from "../../ooxml/precedence";
import {
  attrsOf,
  type Props,
  parseProps,
  propsChild,
  renderProps,
  setChild,
} from "../../ooxml/props";
import {
  HALF_POINTS_PER_PT,
  ST_HpsMeasure,
  ST_SignedTwipsMeasure,
  TWIPS_PER_PT,
} from "../../ooxml/simpleTypes";
import {
  childValue,
  halfPointsToPt,
  isOnElement,
  normalizeHex,
  shadingOf,
  toHexColor,
  twipsToPt,
  wAttr,
} from "../../ooxml/units";
import { childByLocalName } from "../../ooxml/xml";
import { isEastAsianFontName } from "../../styles/fontStack";
import { type ThemeFonts, themeFontName } from "../theme";

/** What text one child of the rPr is to be changed to. A null xml removes that child */
export type ChildEdit = readonly [name: string, xml: string | null];

/** What an edit is written against: the rPr as it stands, and everything the layers below the run lay down */
export interface RunEditContext {
  rPr: Props;
  inherited: RunFormat;
}

type Read<K extends keyof RunFormat> = NonNullable<RunFormat[K]>;

/**
 * The value an edit sets. An off - `false`, an underline of `none` - is not one: it is spelled
 * as null, the withdrawal, and pinned into the rPr only where the layers below would otherwise
 * show through.
 */
export type RunSetting<K extends keyof RunFormat> = Exclude<
  Read<K>,
  false | "none"
>;

/** A run property as it is read and compared. `R` is the value read, `S` the value an edit sets */
export interface RunProperty<R, S> {
  /**
   * The rPr children the property owns, in the order CT_RPr lays them down. A `w:rPrChange`
   * snapshot of the property is exactly these
   */
  readonly children: readonly string[];
  read(rPr: Element, themeFonts: ThemeFonts): R | undefined;
  /** Whether the display values switch or paint the property on, which is what a withdrawal has to undo */
  isOn(format: RunFormat): boolean;
  /** Whether the display values already carry the value, so that writing it would change nothing */
  matches(format: RunFormat, value: S): boolean;
}

/** A run property an edit may write as well */
export interface EditableRunProperty<R, S> extends RunProperty<R, S> {
  /** The children that make the run say the value. null for a value the document cannot record */
  write(value: S, context: RunEditContext): ChildEdit[] | null;
  /** The children that take the run's own say away, pinned off where the layers below would otherwise show through */
  off(context: RunEditContext): ChildEdit[];
}

/** The properties read for display alone. Nothing writes them, so no edit names them */
const DISPLAY_ONLY_RUN_KEYS = ["lang"] as const;
type DisplayOnlyRunKey = (typeof DISPLAY_ONLY_RUN_KEYS)[number];
export type EditableRunKey = Exclude<keyof RunFormat, DisplayOnlyRunKey>;

/**
 * A job that changes one run property. A null value withdraws the run's own setting: the layers
 * below stand again, except where they would switch or paint the property on, in which case the
 * off is pinned into the run.
 */
export type RunEdit<K extends EditableRunKey = EditableRunKey> = {
  [P in K]: { key: P; value: RunSetting<P> | null };
}[K];

/** The settings an rPr built from nothing records */
export type RunSettings = { [K in EditableRunKey]?: RunSetting<K> };

function isEditableRunKey(key: keyof RunFormat): key is EditableRunKey {
  return !DISPLAY_ONLY_RUN_KEYS.some((display) => display === key);
}

/** The editable properties, in the order they are read */
export const EDITABLE_RUN_KEYS: readonly EditableRunKey[] =
  RUN_FORMAT_KEYS.filter(isEditableRunKey);

/** A formatting child that records its whole setting in `w:val`, or one that records nothing */
function valXml(name: string, value: string | null): string {
  return elementXml(wName(name), value === null ? [] : [[wName("val"), value]]);
}

/**
 * The shading that records a text background. `clear` means paint with the `fill` color alone, with no pattern.
 * `fill="auto"` means "no background", so it acts as an off that overrides inheritance.
 */
function shadingXml(fill: string): string {
  return elementXml(wName("shd"), [
    [wName("val"), "clear"],
    [wName("color"), "auto"],
    [wName("fill"), fill],
  ]);
}

/** A toggle property (§17.7.3) as the run wrote it: on, switched off outright, or not mentioned */
function toggleState(rPr: Element, name: string): boolean | null {
  const toggle = childByLocalName(rPr, name);
  return toggle ? isOnElement(toggle) : null;
}

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
  for (const attr of overridingAttrs("rFonts", slot)) {
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

/**
 * A new rFonts with the font name written in.
 *
 * The name goes into the slots `fontSlots` names, and a theme reference in one of those slots
 * is cleared away, since a theme beats a name. The slots the name does not reach keep both
 * their name and their theme reference.
 * The attributes we do not decide (`w:hint` and so on) keep their original values.
 */
function rFontsXml(current: string | null, name: string): string | null {
  const props = current === null ? null : parseProps(current);
  const attrs = props === null ? [] : attrsOf(props);
  if ((current !== null && props === null) || attrs === null) return null;
  const slots = fontSlots(name, wAttrValue(attrs, "cs") !== null);
  const kept = slots.reduce(
    (rest, slot) => setAttr(rest, "rFonts", slot, null),
    attrs
  );
  return elementXml(wName("rFonts"), [
    ...slots.map((slot): XmlAttr => [wName(slot), name]),
    ...kept,
  ]);
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

const VERTICAL_ALIGN_BY_VAL: Record<string, VerticalAlign> = {
  superscript: "superscript",
  subscript: "subscript",
};

type Editable<K extends EditableRunKey> = EditableRunProperty<
  Read<K>,
  RunSetting<K>
>;

/** The children taken away */
function removed(children: readonly string[]): ChildEdit[] {
  return children.map((name) => [name, null]);
}

/**
 * A toggle property (§17.7.3), read from `name` alone.
 * Bold and italic write their complex-script twin beside the Latin one, following the
 * convention of Korean documents; the twin is written, never read.
 */
function toggle<K extends EditableRunKey>(
  key: K,
  name: string,
  ...twins: string[]
): EditableRunProperty<boolean, true> {
  const children = [name, ...twins];
  const isOn = (format: RunFormat): boolean => format[key] === true;
  return {
    children,
    read: (rPr) => toggleState(rPr, name) ?? undefined,
    isOn,
    matches: isOn,
    write: () => children.map((child) => [child, valXml(child, null)]),
    off: ({ inherited }) =>
      isOn(inherited)
        ? children.map((child) => [child, valXml(child, "0")])
        : removed(children),
  };
}

/**
 * An underline holds a kind rather than an on/off state, `none` being the kind that is off.
 * One kind is not told from another: the toggle switches an underline on and leaves the kind a
 * run already has.
 */
const UNDERLINE: Editable<"underline"> = {
  children: ["u"],
  read: (rPr) => {
    const kind = childValue(rPr, "u");
    return isUnderlineKind(kind) || kind === "none" ? kind : undefined;
  },
  isOn: (format) =>
    format.underline !== undefined && format.underline !== "none",
  matches: (format) => UNDERLINE.isOn(format),
  write: (kind) => [["u", valXml("u", kind)]],
  off: ({ inherited }) => [
    ["u", UNDERLINE.isOn(inherited) ? valXml("u", "none") : null],
  ],
};

/** There is no off for a size. Withdrawing the setting inherits the style and the document default again */
const FONT_SIZE: Editable<"fontSizePt"> = {
  children: ["sz", "szCs"],
  read: (rPr) =>
    halfPointsToPt(ST_HpsMeasure.parse(childValue(rPr, "sz"))) ?? undefined,
  isOn: (format) => format.fontSizePt !== undefined,
  matches: (format, pt) => format.fontSizePt === pt,
  write: (pt) => {
    // A size off the half-point step, or past the ceiling, is not one the document can record
    const half = ST_HpsMeasure.format(pt * HALF_POINTS_PER_PT);
    return half === null
      ? null
      : [
          ["sz", valXml("sz", half)],
          ["szCs", valXml("szCs", half)],
        ];
  },
  off: () => removed(FONT_SIZE.children),
};

/** The same holds for the spacing between characters: withdrawn, the layers below decide again */
const CHARACTER_SPACING: Editable<"characterSpacingPt"> = {
  children: ["spacing"],
  read: (rPr) =>
    twipsToPt(ST_SignedTwipsMeasure.parse(childValue(rPr, "spacing"))) ??
    undefined,
  isOn: (format) => format.characterSpacingPt !== undefined,
  matches: (format, pt) => format.characterSpacingPt === pt,
  write: (pt) => {
    const twips = ST_SignedTwipsMeasure.format(Math.round(pt * TWIPS_PER_PT));
    return twips === null ? null : [["spacing", valXml("spacing", twips)]];
  },
  off: () => removed(CHARACTER_SPACING.children),
};

/** There is no off for the font either. Withdrawing the setting removes the whole rFonts so the document default font applies again */
const FONT_FAMILY: Editable<"fontFamily"> = {
  children: ["rFonts"],
  read: (rPr, themeFonts) => {
    const rFonts = childByLocalName(rPr, "rFonts");
    return rFonts ? (fontFamilyOf(rFonts, themeFonts) ?? undefined) : undefined;
  },
  isOn: (format) => fontNamesOf(format.fontFamily).length > 0,
  matches: (format, name) => {
    const names = fontNamesOf(format.fontFamily);
    // If the slots are using different names, it is not yet in the state we want
    return names.length === 1 && names[0] === fontName(name);
  },
  write: (name, { rPr }) => {
    const clean = fontName(name);
    if (clean === null) return null;
    const xml = rFontsXml(
      propsChild(rPr.children, "rFonts")?.xml ?? null,
      clean
    );
    return xml === null ? null : [["rFonts", xml]];
  },
  off: () => removed(FONT_FAMILY.children),
};

/** auto means "the text color as the document decides", so it acts as an off that overrides inheritance */
const COLOR: Editable<"color"> = {
  children: ["color"],
  read: (rPr) => toHexColor(childValue(rPr, "color")) ?? undefined,
  isOn: (format) => format.color !== undefined,
  // Colors differing only in capitalization are the same color
  matches: (format, hex) =>
    format.color !== undefined &&
    normalizeHex(format.color) === normalizeHex(hex),
  write: (hex) => {
    const value = normalizeHex(hex);
    return value === null ? null : [["color", valXml("color", value)]];
  },
  off: ({ inherited }) => [
    ["color", inherited.color === undefined ? null : valXml("color", "auto")],
  ],
};

const HIGHLIGHT: Editable<"highlight"> = {
  children: ["highlight"],
  read: (rPr) => {
    const name = childValue(rPr, "highlight");
    return isHighlightName(name) ? name : undefined;
  },
  isOn: (format) => format.highlight !== undefined,
  matches: (format, name) => format.highlight === name,
  write: (name) => [["highlight", valXml("highlight", name)]],
  off: ({ inherited }) => [
    [
      "highlight",
      inherited.highlight === undefined ? null : valXml("highlight", "none"),
    ],
  ],
};

/**
 * Word paints the highlight on top of the shading. Left in place, it would hide the new
 * background color underneath it, so a background written or withdrawn takes the highlight
 * with it, and text still wearing a highlight is never in the state a background asks for.
 */
const BACKGROUND: Editable<"background"> = {
  children: ["highlight", "shd"],
  read: (rPr) => shadingOf(rPr) ?? undefined,
  isOn: (format) =>
    format.highlight !== undefined ||
    (format.background !== undefined && format.background !== NO_FILL),
  matches: (format, hex) => {
    const fill = normalizeHex(hex);
    return (
      fill !== null &&
      format.highlight === undefined &&
      format.background !== undefined &&
      normalizeHex(format.background) === fill
    );
  },
  write: (hex, context) => {
    const fill = normalizeHex(hex);
    return fill === null
      ? null
      : [["shd", shadingXml(fill)], ...HIGHLIGHT.off(context)];
  },
  off: (context) => {
    const painted =
      context.inherited.background !== undefined &&
      context.inherited.background !== NO_FILL;
    return [
      ["shd", painted ? shadingXml("auto") : null],
      ...HIGHLIGHT.off(context),
    ];
  },
};

/** `baseline` is the position the document decides, so it acts as an off that overrides inheritance */
const VERTICAL_ALIGN: Editable<"verticalAlign"> = {
  children: ["vertAlign"],
  read: (rPr) => VERTICAL_ALIGN_BY_VAL[childValue(rPr, "vertAlign") ?? ""],
  isOn: (format) => format.verticalAlign !== undefined,
  matches: (format, align) => format.verticalAlign === align,
  write: (align) => [["vertAlign", valXml("vertAlign", align)]],
  off: ({ inherited }) => [
    [
      "vertAlign",
      inherited.verticalAlign === undefined
        ? null
        : valXml("vertAlign", "baseline"),
    ],
  ],
};

const LANG: RunProperty<Read<"lang">, RunSetting<"lang">> = {
  children: ["lang"],
  read: (rPr) => langOf(rPr) ?? undefined,
  isOn: (format) => format.lang !== undefined,
  matches: (format, tag) => format.lang === tag,
};

const EDITABLE_RUN_PROPERTIES: { [K in EditableRunKey]: Editable<K> } = {
  bold: toggle("bold", "b", "bCs"),
  italic: toggle("italic", "i", "iCs"),
  strike: toggle("strike", "strike"),
  doubleStrike: toggle("doubleStrike", "dstrike"),
  smallCaps: toggle("smallCaps", "smallCaps"),
  caps: toggle("caps", "caps"),
  underline: UNDERLINE,
  fontSizePt: FONT_SIZE,
  characterSpacingPt: CHARACTER_SPACING,
  fontFamily: FONT_FAMILY,
  color: COLOR,
  highlight: HIGHLIGHT,
  background: BACKGROUND,
  verticalAlign: VERTICAL_ALIGN,
};

const DISPLAY_ONLY_RUN_PROPERTIES: {
  [K in DisplayOnlyRunKey]: RunProperty<Read<K>, RunSetting<K>>;
} = { lang: LANG };

export const RUN_PROPERTIES: {
  [K in keyof Required<RunFormat>]: RunProperty<Read<K>, RunSetting<K>>;
} = { ...EDITABLE_RUN_PROPERTIES, ...DISPLAY_ONLY_RUN_PROPERTIES };

/** Which children of the rPr one job changes and how. null for a value whose meaning cannot be made out */
export function runChildEdits<K extends EditableRunKey>(
  edit: RunEdit<K>,
  context: RunEditContext
): ChildEdit[] | null {
  const property = EDITABLE_RUN_PROPERTIES[edit.key];
  return edit.value === null
    ? property.off(context)
    : property.write(edit.value, context);
}

/** Whether this text is already in the state the job wants. If it already is, we leave it untouched */
export function matchesRunEdit<K extends EditableRunKey>(
  format: RunFormat | null,
  edit: RunEdit<K>
): boolean {
  const property = RUN_PROPERTIES[edit.key];
  const values = format ?? {};
  return edit.value === null
    ? !property.isOn(values)
    : property.matches(values, edit.value);
}

/** The rPr with these children swapped in */
export function withChildEdits(rPr: Props, edits: readonly ChildEdit[]): Props {
  return edits.reduce((kept, [name, xml]) => setChild(kept, name, xml), rPr);
}

/** Turning formatting on for a run with no formatting creates a minimal rPr from scratch */
export const EMPTY_RUN_PROPS: Props = {
  tag: wName("rPr"),
  attrs: null,
  children: [],
};

function written<K extends EditableRunKey>(
  rPr: Props,
  key: K,
  value: RunSetting<K> | undefined
): Props {
  if (value === undefined) return rPr;
  const edits = EDITABLE_RUN_PROPERTIES[key].write(value, {
    rPr,
    inherited: {},
  });
  return edits === null ? rPr : withChildEdits(rPr, edits);
}

/**
 * A whole rPr recording these settings, built from nothing, and null where they record nothing.
 * A setting the document cannot record is left out, the way an edit refusing it leaves the run
 * as it was.
 */
export function rPrOf(settings: RunSettings): string | null {
  const rPr = renderProps(
    EDITABLE_RUN_KEYS.reduce(
      (kept, key) => written(kept, key, settings[key]),
      EMPTY_RUN_PROPS
    )
  );
  return rPr === "" ? null : rPr;
}
