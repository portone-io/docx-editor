/**
 * Reads numbering.xml to build up "what shape does level N of list M have".
 *
 * The values produced here are used only for drawing the numbers on screen.
 * numbering.xml itself goes out as the original bytes on export, so this reading never
 * disturbs the file.
 */

import type { TabStopDirective } from "../model/tabStops";
import {
  ST_DecimalNumber,
  ST_SignedTwipsMeasure,
  ST_TwipsMeasure,
} from "../ooxml/simpleTypes";
import { readTabStopDirectives } from "../ooxml/tabStops";
import {
  ALIGN_BY_JC,
  childValue,
  isOn,
  twipsToPt,
  wAttr,
} from "../ooxml/units";
import {
  childByLocalName,
  elementChildren,
  parseXml,
  withXmlParser,
  type XmlParser,
} from "../ooxml/xml";

/** The number formats actually used across every fixture */
const NUMBER_FORMATS = [
  "decimal",
  "bullet",
  "lowerLetter",
  "upperLetter",
  "lowerRoman",
] as const;

export type NumberFormat = (typeof NUMBER_FORMATS)[number];

/**
 * The indent a level defines. The unit is twips (1/20 of a point), and a slot that is
 * not recorded is null
 */
export interface LevelIndent {
  startTwips: number | null;
  endTwips: number | null;
  hangingTwips: number | null;
  firstLineTwips: number | null;
}

/** What stands between a level's number and the text of the paragraph (`w:suff`, §17.9.28) */
export type LevelSuffix = "tab" | "space" | "nothing";

/** Where a level's number sits in the space kept for it (`w:lvlJc`, §17.9.7) */
export type LevelAlign = "left" | "center" | "right";

export interface NumberingLevel {
  format: NumberFormat;
  /**
   * The pattern holding the number slots, such as `%1.` or `(%2)`. For a bullet it is
   * the bullet glyph itself
   */
  text: string;
  start: number;
  /** The indent the paragraphs of this level use. Null when the level defines none */
  indent: LevelIndent | null;
  /** Custom stops contributed by this level's paragraph properties. */
  tabStops?: readonly TabStopDirective[];
  /**
   * The deepest level whose advance sends this one back to its start (`w:lvlRestart`, §17.9.10),
   * as a level number: any level above that one restarts it too. A level of -1 never restarts, and
   * null is the default, the level right above this one.
   */
  restartAfterLevel: number | null;
  /** Every level in this level's text spelled as a decimal, whatever format it counts in (`w:isLgl`, §17.9.4) */
  legal: boolean;
  suffix: LevelSuffix;
  align: LevelAlign;
}

export interface NumberingList {
  /** The shape for each level (ilvl) */
  levels: Map<number, NumberingLevel>;
}

export interface Numbering {
  /** The list for each numId */
  lists: Map<number, NumberingList>;
}

export const EMPTY_NUMBERING: Numbering = { lists: new Map() };

function isNumberFormat(value: string | null): value is NumberFormat {
  return NUMBER_FORMATS.some((format) => format === value);
}

function indentOf(lvl: Element): LevelIndent | null {
  const pPr = childByLocalName(lvl, "pPr");
  const ind = pPr ? childByLocalName(pPr, "ind") : null;
  if (!ind) return null;
  const startTwips = ST_SignedTwipsMeasure.parse(
    wAttr(ind, "start") ?? wAttr(ind, "left")
  );
  const endTwips = ST_SignedTwipsMeasure.parse(
    wAttr(ind, "end") ?? wAttr(ind, "right")
  );
  const indent: LevelIndent = {
    startTwips,
    endTwips,
    hangingTwips: ST_TwipsMeasure.parse(wAttr(ind, "hanging")),
    firstLineTwips: ST_TwipsMeasure.parse(wAttr(ind, "firstLine")),
  };
  const empty =
    startTwips === null &&
    endTwips === null &&
    indent.hangingTwips === null &&
    indent.firstLineTwips === null;
  return empty ? null : indent;
}

/**
 * The indent a level defines, converted into the on-screen unit (points). A slot the
 * level does not define is null
 */
export interface LevelIndentPt {
  startPt: number | null;
  endPt: number | null;
  /**
   * Negative for a hanging indent, positive for a first-line indent. Null when there is
   * neither
   */
  textIndentPt: number | null;
}

export const NO_LEVEL_INDENT: LevelIndentPt = {
  startPt: null,
  endPt: null,
  textIndentPt: null,
};

/**
 * The indent a level passes down to its paragraphs.
 * As with paragraph formatting, a hanging indent overrides a first-line indent (the
 * OOXML rule).
 */
export function levelIndentPt(indent: LevelIndent | null): LevelIndentPt {
  if (!indent) return NO_LEVEL_INDENT;
  const hangingPt = twipsToPt(indent.hangingTwips);
  return {
    startPt: twipsToPt(indent.startTwips),
    endPt: twipsToPt(indent.endTwips),
    textIndentPt:
      hangingPt !== null ? -hangingPt : twipsToPt(indent.firstLineTwips),
  };
}

const LEVEL_SUFFIXES: readonly LevelSuffix[] = ["tab", "space", "nothing"];

/** A level that says nothing puts a tab between its number and the text (§17.9.28) */
function suffixOf(lvl: Element): LevelSuffix {
  const suffix = childValue(lvl, "suff");
  return LEVEL_SUFFIXES.find((known) => known === suffix) ?? "tab";
}

/**
 * Where the number sits in the space kept for it.
 * A level that says nothing, or that asks for a justification a number cannot take, keeps it at
 * the left, which is what §17.9.7 gives a left-to-right paragraph.
 */
function alignOf(lvl: Element): LevelAlign {
  const jc = childValue(lvl, "lvlJc");
  const align = jc === null ? undefined : ALIGN_BY_JC[jc];
  return align === undefined || align === "justify" ? "left" : align;
}

/**
 * Which level's advance restarts this one.
 *
 * `w:lvlRestart` counts levels from one, so the level it names is the number minus one, and the 0
 * that §17.9.10 gives for a level that never restarts lands below the outermost level of all. A
 * number that cannot be read at all leaves the default standing.
 */
function restartAfterLevelOf(lvl: Element): number | null {
  const restart = ST_DecimalNumber.parse(childValue(lvl, "lvlRestart"));
  return restart === null ? null : restart - 1;
}

function readLevel(lvl: Element): NumberingLevel {
  const format = childValue(lvl, "numFmt");
  const pPr = childByLocalName(lvl, "pPr");
  const tabStops = readTabStopDirectives(pPr);
  return {
    format: isNumberFormat(format) ? format : "decimal",
    text: childValue(lvl, "lvlText") ?? "",
    start: ST_DecimalNumber.parse(childValue(lvl, "start")) ?? 1,
    indent: indentOf(lvl),
    ...(tabStops.length === 0 ? {} : { tabStops }),
    restartAfterLevel: restartAfterLevelOf(lvl),
    legal: isOn(lvl, "isLgl"),
    suffix: suffixOf(lvl),
    align: alignOf(lvl),
  };
}

/** Collects each `<w:lvl w:ilvl="0">` bundle, keyed by its level number */
function readLevels(parent: Element): Map<number, NumberingLevel> {
  const levels = new Map<number, NumberingLevel>();
  for (const child of elementChildren(parent)) {
    if (child.localName !== "lvl") continue;
    const ilvl = ST_DecimalNumber.parse(wAttr(child, "ilvl"));
    if (ilvl !== null) levels.set(ilvl, readLevel(child));
  }
  return levels;
}

/**
 * Takes the abstractNum that `<w:num>` refers to as the outline and overrides it wherever
 * an lvlOverride is present.
 * An lvlOverride may swap out a whole level or change only its start number.
 */
function readList(
  num: Element,
  abstractLevels: Map<number, Map<number, NumberingLevel>>
): NumberingList | null {
  const abstractNumId = ST_DecimalNumber.parse(
    childValue(num, "abstractNumId")
  );
  if (abstractNumId === null) return null;
  const levels = new Map(abstractLevels.get(abstractNumId) ?? []);

  for (const child of elementChildren(num)) {
    if (child.localName !== "lvlOverride") continue;
    const ilvl = ST_DecimalNumber.parse(wAttr(child, "ilvl"));
    if (ilvl === null) continue;
    const replacement = childByLocalName(child, "lvl");
    const base = replacement ? readLevel(replacement) : levels.get(ilvl);
    if (!base) continue;
    const startOverride = ST_DecimalNumber.parse(
      childValue(child, "startOverride")
    );
    levels.set(
      ilvl,
      startOverride === null ? base : { ...base, start: startOverride }
    );
  }
  return { levels };
}

/**
 * The list each numbering style names, by the style's id.
 *
 * A numbering style is never worn by a paragraph: it is a name on a list, which an abstract
 * definition reaches through `w:numStyleLink` (§17.9.21). The map is built from styles.xml, which
 * this folder cannot read for itself.
 */
export type NumberingStyleLinks = ReadonlyMap<string, number>;

/** What a caller may say about reading numbering beyond handing over the XML */
export interface NumberingOptions {
  /**
   * The parser the numbering XML is read through. Left out, the `DOMParser` global is used, and a
   * runtime carrying none refuses the read with `no-xml-parser`.
   */
  xmlParser?: XmlParser;
  /**
   * The list each numbering style of the document names. Left out, a definition deferring to a
   * numbering style is read through the definition that declares it stands behind that style.
   */
  links?: NumberingStyleLinks;
}

const NO_STYLE_LINKS: NumberingStyleLinks = new Map();

/** What the abstract definitions of one numbering part say about each other */
interface Definitions {
  /** Each `w:abstractNum` by its id */
  byId: ReadonlyMap<number, Element>;
  /** The definition each `w:num` names */
  ofList: ReadonlyMap<number, number>;
  /** The definition that declares itself the one behind a numbering style (`w:styleLink`) */
  ofStyle: ReadonlyMap<string, number>;
  links: NumberingStyleLinks;
}

/**
 * The levels one abstract definition lays down.
 *
 * A definition carrying `w:numStyleLink` holds none of its own and defers to a numbering style
 * (§17.9.21), which names the list whose definition holds them. That definition is the one the
 * style points at, and where no style table was handed in, the one that declares it stands behind
 * that style (`w:styleLink`, §17.9.27). A link that leads back to a definition already followed is
 * left where it is rather than followed round again.
 */
function levelsOf(
  id: number,
  definitions: Definitions,
  seen: Set<number>
): Map<number, NumberingLevel> {
  const el = definitions.byId.get(id);
  if (!el) return new Map();
  const own = readLevels(el);
  const styleId = childValue(el, "numStyleLink");
  if (own.size > 0 || styleId === null) return own;

  const named = definitions.links.get(styleId);
  const deferred =
    (named === undefined ? undefined : definitions.ofList.get(named)) ??
    definitions.ofStyle.get(styleId);
  if (deferred === undefined || seen.has(deferred)) return own;
  seen.add(deferred);
  return levelsOf(deferred, definitions, seen);
}

/**
 * The lists this numbering XML defines. `null` is the document that has no numbering part, which
 * defines none and asks nothing of a parser.
 */
export function parseNumbering(
  xml: string | null,
  options?: NumberingOptions
): Numbering {
  if (xml === null) return EMPTY_NUMBERING;
  return withXmlParser(options?.xmlParser, () => readNumbering(xml, options));
}

function readNumbering(xml: string, options?: NumberingOptions): Numbering {
  const root = parseXml(xml).documentElement;

  const byId = new Map<number, Element>();
  const ofStyle = new Map<string, number>();
  const ofList = new Map<number, number>();
  for (const child of elementChildren(root)) {
    if (child.localName === "abstractNum") {
      const id = ST_DecimalNumber.parse(wAttr(child, "abstractNumId"));
      if (id === null) continue;
      byId.set(id, child);
      const styleId = childValue(child, "styleLink");
      if (styleId !== null) ofStyle.set(styleId, id);
    } else if (child.localName === "num") {
      const numId = ST_DecimalNumber.parse(wAttr(child, "numId"));
      const id = ST_DecimalNumber.parse(childValue(child, "abstractNumId"));
      if (numId !== null && id !== null) ofList.set(numId, id);
    }
  }
  const definitions: Definitions = {
    byId,
    ofList,
    ofStyle,
    links: options?.links ?? NO_STYLE_LINKS,
  };

  const abstractLevels = new Map<number, Map<number, NumberingLevel>>();
  for (const id of byId.keys()) {
    abstractLevels.set(id, levelsOf(id, definitions, new Set([id])));
  }

  const lists = new Map<number, NumberingList>();
  for (const child of elementChildren(root)) {
    if (child.localName !== "num") continue;
    const numId = ST_DecimalNumber.parse(wAttr(child, "numId"));
    const list = numId === null ? null : readList(child, abstractLevels);
    if (numId !== null && list) lists.set(numId, list);
  }
  return { lists };
}
