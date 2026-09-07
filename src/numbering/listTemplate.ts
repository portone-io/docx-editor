/**
 * The standard shape of a newly started list.
 *
 * The definition of a new list is not kept in the document model.
 * So its shape has to be derivable from the list number (numId) alone, and that rule
 * lives here. Even numbers are numbered lists, odd numbers are bullet lists.
 *
 * Drawing the markers on screen and writing the definitions into numbering.xml both look
 * at the same template.
 */

import { elementXml, type XmlAttr } from "../ooxml/element";
import { wName } from "../ooxml/names";
import { orderedElement } from "../ooxml/props";
import type {
  LevelIndent,
  NumberFormat,
  Numbering,
  NumberingLevel,
  NumberingList,
} from "./parseNumbering";

export type ListKind = "numbered" | "bullet";

/** A list has nine levels (the count OOXML defines) */
export const LEVEL_COUNT = 9;

export const MAX_ILVL = LEVEL_COUNT - 1;

/**
 * The number formats cycle through 1. / a. / i. (the default in Word and in the contract
 * formats)
 */
const NUMBER_CYCLE: readonly NumberFormat[] = [
  "decimal",
  "lowerLetter",
  "lowerRoman",
];

const BULLET_CYCLE: readonly string[] = ["●", "○", "■"];

/** Each level indents a further 0.5 inch (720 twips), and the marker hangs out by 360 twips */
const LEVEL_STEP_TWIPS = 720;
const HANGING_TWIPS = 360;

export function templateIndent(ilvl: number): LevelIndent {
  return {
    startTwips: LEVEL_STEP_TWIPS * (ilvl + 1),
    endTwips: null,
    hangingTwips: HANGING_TWIPS,
    firstLineTwips: null,
  };
}

function cycled<T>(cycle: readonly T[], ilvl: number, fallback: T): T {
  return cycle[ilvl % cycle.length] ?? fallback;
}

function templateLevel(kind: ListKind, ilvl: number): NumberingLevel {
  const numbered = kind === "numbered";
  return {
    format: numbered ? cycled(NUMBER_CYCLE, ilvl, "decimal") : "bullet",
    text: numbered ? `%${ilvl + 1}.` : cycled(BULLET_CYCLE, ilvl, "●"),
    start: 1,
    indent: templateIndent(ilvl),
    restartAfterLevel: null,
    legal: false,
    suffix: "tab",
    align: "left",
  };
}

function templateList(kind: ListKind): NumberingList {
  const levels = new Map<number, NumberingLevel>();
  for (let ilvl = 0; ilvl < LEVEL_COUNT; ilvl += 1) {
    levels.set(ilvl, templateLevel(kind, ilvl));
  }
  return { levels };
}

const TEMPLATES: Record<ListKind, NumberingList> = {
  numbered: templateList("numbered"),
  bullet: templateList("bullet"),
};

export function listKindOf(numId: number): ListKind {
  return numId % 2 === 0 ? "numbered" : "bullet";
}

/**
 * The list definition for this number.
 * A number the document's numbering.xml does not know is read as the standard template
 * (which is the case for a list that was just started).
 */
export function listFor(numbering: Numbering, numId: number): NumberingList {
  return numbering.lists.get(numId) ?? TEMPLATES[listKindOf(numId)];
}

/** The next number not yet in use. Picks the even or odd one matching the shape */
export function nextNumId(used: Iterable<number>, kind: ListKind): number {
  const from = Math.max(0, ...used) + 1;
  return listKindOf(from) === kind ? from : from + 1;
}

/** Writes only the slots that carry a value. If none do, no indent is written at all */
function indXml(indent: LevelIndent | null): string {
  const slots: readonly (readonly [name: string, twips: number | null])[] = [
    ["left", indent?.startTwips ?? null],
    ["right", indent?.endTwips ?? null],
    ["hanging", indent?.hangingTwips ?? null],
    ["firstLine", indent?.firstLineTwips ?? null],
  ];
  const attrs = slots
    .filter((slot): slot is readonly [string, number] => slot[1] !== null)
    .map(([name, twips]): XmlAttr => [wName(name), `${twips}`]);
  if (attrs.length === 0) return "";
  return orderedElement(
    wName("pPr"),
    [],
    [{ name: "ind", xml: elementXml(wName("ind"), attrs) }]
  );
}

function levelXml(ilvl: number, level: NumberingLevel): string {
  const ind = indXml(level.indent);
  return orderedElement(
    wName("lvl"),
    [[wName("ilvl"), `${ilvl}`]],
    [
      {
        name: "start",
        xml: elementXml(wName("start"), [[wName("val"), `${level.start}`]]),
      },
      {
        name: "numFmt",
        xml: elementXml(wName("numFmt"), [[wName("val"), level.format]]),
      },
      {
        name: "lvlText",
        xml: elementXml(wName("lvlText"), [[wName("val"), level.text]]),
      },
      {
        name: "lvlJc",
        xml: elementXml(wName("lvlJc"), [[wName("val"), "left"]]),
      },
      ...(ind === "" ? [] : [{ name: "pPr", xml: ind }]),
    ]
  );
}

/** Writes one template out as a `<w:abstractNum>` definition */
export function abstractNumXml(abstractNumId: number, kind: ListKind): string {
  const levels = TEMPLATES[kind].levels;
  const body = Array.from(levels.entries())
    .map(([ilvl, level]) => levelXml(ilvl, level))
    .join("");
  return elementXml(
    wName("abstractNum"),
    [[wName("abstractNumId"), `${abstractNumId}`]],
    [body]
  );
}

/** The `<w:num>` that ties a list number to a definition */
export function numXml(numId: number, abstractNumId: number): string {
  return elementXml(
    wName("num"),
    [[wName("numId"), `${numId}`]],
    [elementXml(wName("abstractNumId"), [[wName("val"), `${abstractNumId}`]])]
  );
}
