/** Reads and layers styles.xml values used by the display model. */

import {
  type CellMargins,
  type InsideBorders,
  type ParagraphFormat,
  type RunFormat,
  type TableFormat,
  toParagraphFormat,
  toRunFormat,
  toTableFormat,
} from "../../model/format";
import { childValue, isOn, wAttr } from "../../ooxml/units";
import { childByLocalName, elementChildren } from "../../ooxml/xml";
import { parsePropsXml } from "../propsXml";
import {
  layerCellMargins,
  layerInsideBorders,
  NO_CELL_MARGINS,
  NO_INSIDE_BORDERS,
  readCellMarginsOf,
  readInsideBorders,
  readTableFormat,
} from "../tableFormatting";
import { NO_THEME_FONTS, type ThemeFonts } from "../theme";
import { readParagraphFormat, readRunFormat } from "./direct";
import { layerParagraphValues, type ParagraphFormatLayer } from "./tabStops";

/** The display values one style passes down to paragraphs, to text, and to tables */
export interface StyleFormat {
  paragraph: ParagraphFormatLayer;
  run: RunFormat;
  table: TableFormat;
  /** The lines between cells the style laid down. A side it says nothing about is null */
  tableInside: InsideBorders;
  /** The cell margins the style laid down. A side it says nothing about is null */
  tableCellMargins: CellMargins;
}

/** A table looked up by style name. The `basedOn` chain has already been layered into the values */
export type StyleTable = ReadonlyMap<string, StyleFormat>;

export const NO_STYLES: StyleTable = new Map();

/** One style as styles.xml wrote it down. This is its shape before the chain is layered in */
interface StyleSource {
  basedOn: string | null;
  pPr: Element | null;
  rPr: Element | null;
  tblPr: Element | null;
}

/**
 * The chain from the root down to this style.
 * If the chain points back at itself or at a style that does not exist, it stops right there.
 */
function styleChain(
  id: string,
  sources: Map<string, StyleSource>
): StyleSource[] {
  const chain: StyleSource[] = [];
  const seen = new Set<string>();
  let current: string | null = id;
  while (current !== null && !seen.has(current)) {
    seen.add(current);
    const source = sources.get(current);
    if (!source) break;
    chain.unshift(source);
    current = source.basedOn;
  }
  return chain;
}

const EMPTY_STYLE_FORMAT: StyleFormat = {
  paragraph: {},
  run: {},
  table: {},
  tableInside: NO_INSIDE_BORDERS,
  tableCellMargins: NO_CELL_MARGINS,
};

/** The values one style wrote down itself, before the styles it is based on are layered underneath */
function readStyleFormat(
  source: StyleSource,
  themeFonts: ThemeFonts
): StyleFormat {
  return {
    paragraph: readParagraphFormat(source.pPr) ?? {},
    run: readRunFormat(source.rPr, themeFonts) ?? {},
    table: readTableFormat(source.tblPr) ?? {},
    tableInside: readInsideBorders(source.tblPr),
    tableCellMargins: readCellMarginsOf(source.tblPr, "tblCellMar"),
  };
}

/**
 * Lays the values of one style on top of the ones its ancestors laid down.
 *
 * The inside lines are layered one side at a time, because a style that says nothing about a side
 * has to leave the ancestor's line standing. A style that writes the side as `none` or `nil` reads
 * back as the string "none", so it switches the ancestor's line off rather than leaving it alone.
 * The cell margins are layered the same way, side by side.
 */
function layerStyleFormat(base: StyleFormat, over: StyleFormat): StyleFormat {
  return {
    paragraph: layerParagraphValues(base.paragraph, over.paragraph),
    run: { ...base.run, ...over.run },
    table: { ...base.table, ...over.table },
    tableInside: layerInsideBorders(base.tableInside, over.tableInside),
    tableCellMargins: layerCellMargins(
      base.tableCellMargins,
      over.tableCellMargins
    ),
  };
}

/** The effective values layered down from the root. Further down the chain overrides further up */
function foldChain(chain: StyleSource[], themeFonts: ThemeFonts): StyleFormat {
  return chain
    .map((source) => readStyleFormat(source, themeFonts))
    .reduce(layerStyleFormat, EMPTY_STYLE_FORMAT);
}

/**
 * Reads the style chain from styles.xml into effective values for display.
 *
 * A style name is unique across all kinds (paragraph, character, table), so we take them in
 * without telling the kinds apart.
 * The names a paragraph points at are only ever paragraph styles, so the other kinds are never looked up.
 */
export function readStyles(
  styles: Document,
  themeFonts: ThemeFonts = NO_THEME_FONTS
): StyleTable {
  const sources = new Map<string, StyleSource>();
  for (const el of elementChildren(styles.documentElement)) {
    if (el.localName !== "style") continue;
    const id = wAttr(el, "styleId");
    if (id === null || id.length === 0) continue;
    sources.set(id, {
      basedOn: childValue(el, "basedOn"),
      pPr: childByLocalName(el, "pPr"),
      rPr: childByLocalName(el, "rPr"),
      tblPr: childByLocalName(el, "tblPr"),
    });
  }

  const table = new Map<string, StyleFormat>();
  for (const id of sources.keys()) {
    table.set(id, foldChain(styleChain(id, sources), themeFonts));
  }
  return table;
}

/** Whether the style is the one OOXML applies to every object of its kind that points at no style */
function isDefaultStyle(el: Element): boolean {
  const value = wAttr(el, "default");
  return value === "1" || value === "true";
}

/**
 * The id of the style the document marked as its default (`w:default="1"`) for one kind of object.
 *
 * OOXML applies a default style to every object of that kind that points at no style of its own.
 * When more than one style claims to be the default, the last one wins, as Word reads it.
 */
function defaultStyleIdOf(styles: Document, type: string): string | null {
  let found: string | null = null;
  for (const el of elementChildren(styles.documentElement)) {
    if (el.localName !== "style" || wAttr(el, "type") !== type) continue;
    if (!isDefaultStyle(el)) continue;
    const id = wAttr(el, "styleId");
    if (id !== null && id.length > 0) found = id;
  }
  return found;
}

/**
 * The default table style, which for most documents is where a plain table gets its cell margins
 * and its lines from
 */
export function defaultTableStyleIdOf(styles: Document): string | null {
  return defaultStyleIdOf(styles, "table");
}

/** The default paragraph style (usually Normal), which every paragraph with no `w:pStyle` wears */
export function defaultParagraphStyleIdOf(styles: Document): string | null {
  return defaultStyleIdOf(styles, "paragraph");
}

/** One paragraph style the document defines, as the style picker shows it */
export interface ParagraphStyleOption {
  id: string;
  /** The display name `w:name` states. Falls back on the id */
  name: string;
  /** Set on the one style the document marked `w:default="1"`, which a paragraph with no pStyle already wears */
  isDefault: boolean;
  /** Set on a primary style (`w:qFormat`) intended for a prominent style gallery */
  primary: boolean;
  /** Set on a style Word keeps out of its own gallery */
  hidden: boolean;
}

/**
 * The paragraph styles a document defines, in the order styles.xml lists them.
 *
 * The latent styles a document carries around without using them (`w:semiHidden`, `w:hidden`) come
 * along all the same, flagged as hidden, because a paragraph can still point at one and then its
 * `w:name` is the only name there is to show for it.
 * Where several styles claim to be the default, only the last one is marked, the same rule
 * `defaultParagraphStyleIdOf` reads them by.
 */
export function readParagraphStyles(styles: Document): ParagraphStyleOption[] {
  const defaultId = defaultParagraphStyleIdOf(styles);
  const options: ParagraphStyleOption[] = [];
  for (const el of elementChildren(styles.documentElement)) {
    if (el.localName !== "style" || wAttr(el, "type") !== "paragraph") continue;
    const id = wAttr(el, "styleId");
    if (id === null || id.length === 0) continue;
    const name = childValue(el, "name");
    options.push({
      id,
      name: name === null || name.length === 0 ? id : name,
      isDefault: id === defaultId,
      primary: isOn(el, "qFormat"),
      hidden: isOn(el, "semiHidden") || isOn(el, "hidden"),
    });
  }
  return options;
}

/**
 * The style ids already read out of a formatting fragment.
 *
 * The toolbar decides the style of every selected paragraph on each render, and parsing the same
 * fragment over and over is the whole cost of that. The fragment text is the key, so one fragment
 * is parsed once however many paragraphs share it.
 */
const styleIdsByPPr = new Map<string, string | null>();

/** Past this many fragments the table is dropped rather than grown for as long as the page lives */
const STYLE_ID_CACHE_LIMIT = 2000;

/** The style name the paragraph formatting XML points at. null if it points at none */
export function styleIdOf(pPr: unknown): string | null {
  // Paragraphs that point at a style are rare, so we screen them out first with a cheap string check
  if (typeof pPr !== "string" || !pPr.includes("pStyle")) return null;
  const known = styleIdsByPPr.get(pPr);
  if (known !== undefined) return known;
  const el = parsePropsXml(pPr);
  const id = el ? childValue(el, "pStyle") : null;
  if (styleIdsByPPr.size >= STYLE_ID_CACHE_LIMIT) styleIdsByPPr.clear();
  styleIdsByPPr.set(pPr, id);
  return id;
}

/**
 * The values the style a paragraph wears lays down.
 *
 * That is the style its `w:pStyle` names, and for a paragraph naming none, the document's default
 * paragraph style: OOXML applies that one to every paragraph that points at no style of its own.
 */
export function paragraphStyleFormat(
  pPr: unknown,
  styles: StyleTable,
  defaultStyleId: string | null = null
): StyleFormat | undefined {
  const id = styleIdOf(pPr) ?? defaultStyleId;
  return id === null ? undefined : styles.get(id);
}

/** When there is neither a value from the style nor direct formatting, the display value is left absent too */
function layered(
  style: object,
  direct: object | null
): Record<string, unknown> | null {
  if (direct === null && Object.keys(style).length === 0) return null;
  return { ...style, ...direct };
}

/** Lays direct formatting on top of the values the style laid down. What the paragraph wrote down always wins */
export function layerParagraphFormat(
  style: ParagraphFormatLayer,
  direct: ParagraphFormatLayer | null
): ParagraphFormat | null {
  if (direct === null && Object.keys(style).length === 0) return null;
  return toParagraphFormat(layerParagraphValues(style, direct ?? {}));
}

export function layerRunFormat(
  style: RunFormat,
  direct: RunFormat | null
): RunFormat | null {
  return toRunFormat(layered(style, direct));
}

export function layerTableFormat(
  style: TableFormat,
  direct: TableFormat | null
): TableFormat | null {
  return toTableFormat(layered(style, direct));
}
