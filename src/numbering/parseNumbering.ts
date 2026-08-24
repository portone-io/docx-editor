/**
 * Reads numbering.xml to build up "what shape does level N of list M have".
 *
 * The values produced here are used only for drawing the numbers on screen.
 * numbering.xml itself goes out as the original bytes on export, so this reading never
 * disturbs the file.
 */

import type { TabStopDirective } from "../model/tabStops";
import { readTabStopDirectives } from "../ooxml/tabStops";
import { childValue, round, wAttr } from "../ooxml/units";
import { childByLocalName, elementChildren, parseXml } from "../ooxml/xml";

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

function toInteger(value: string | null): number | null {
  if (value === null) return null;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function isNumberFormat(value: string | null): value is NumberFormat {
  return NUMBER_FORMATS.some((format) => format === value);
}

function indentOf(lvl: Element): LevelIndent | null {
  const pPr = childByLocalName(lvl, "pPr");
  const ind = pPr ? childByLocalName(pPr, "ind") : null;
  if (!ind) return null;
  const startTwips = toInteger(wAttr(ind, "start") ?? wAttr(ind, "left"));
  const endTwips = toInteger(wAttr(ind, "end") ?? wAttr(ind, "right"));
  const indent: LevelIndent = {
    startTwips,
    endTwips,
    hangingTwips: toInteger(wAttr(ind, "hanging")),
    firstLineTwips: toInteger(wAttr(ind, "firstLine")),
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

function twipsToPt(twips: number | null): number | null {
  return twips === null ? null : round(twips / 20);
}

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

function readLevel(lvl: Element): NumberingLevel {
  const format = childValue(lvl, "numFmt");
  const pPr = childByLocalName(lvl, "pPr");
  const tabStops = readTabStopDirectives(pPr);
  return {
    format: isNumberFormat(format) ? format : "decimal",
    text: childValue(lvl, "lvlText") ?? "",
    start: toInteger(childValue(lvl, "start")) ?? 1,
    indent: indentOf(lvl),
    ...(tabStops.length === 0 ? {} : { tabStops }),
  };
}

/** Collects each `<w:lvl w:ilvl="0">` bundle, keyed by its level number */
function readLevels(parent: Element): Map<number, NumberingLevel> {
  const levels = new Map<number, NumberingLevel>();
  for (const child of elementChildren(parent)) {
    if (child.localName !== "lvl") continue;
    const ilvl = toInteger(wAttr(child, "ilvl"));
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
  const abstractNumId = toInteger(childValue(num, "abstractNumId"));
  if (abstractNumId === null) return null;
  const levels = new Map(abstractLevels.get(abstractNumId) ?? []);

  for (const child of elementChildren(num)) {
    if (child.localName !== "lvlOverride") continue;
    const ilvl = toInteger(wAttr(child, "ilvl"));
    if (ilvl === null) continue;
    const replacement = childByLocalName(child, "lvl");
    const base = replacement ? readLevel(replacement) : levels.get(ilvl);
    if (!base) continue;
    const startOverride = toInteger(childValue(child, "startOverride"));
    levels.set(
      ilvl,
      startOverride === null ? base : { ...base, start: startOverride }
    );
  }
  return { levels };
}

export function parseNumbering(xml: string | null): Numbering {
  if (xml === null) return EMPTY_NUMBERING;
  const root = parseXml(xml).documentElement;

  const abstractLevels = new Map<number, Map<number, NumberingLevel>>();
  for (const child of elementChildren(root)) {
    if (child.localName !== "abstractNum") continue;
    const id = toInteger(wAttr(child, "abstractNumId"));
    if (id !== null) abstractLevels.set(id, readLevels(child));
  }

  const lists = new Map<number, NumberingList>();
  for (const child of elementChildren(root)) {
    if (child.localName !== "num") continue;
    const numId = toInteger(wAttr(child, "numId"));
    const list = numId === null ? null : readList(child, abstractLevels);
    if (numId !== null && list) lists.set(numId, list);
  }
  return { lists };
}
