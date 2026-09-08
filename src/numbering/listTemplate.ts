/**
 * A list started while editing: the definition it is started with, the number it takes, and the
 * `<w:abstractNum>` that definition is written out as.
 *
 * Nothing here reads anything into the number. The definition is a value the editor registers
 * (`./listRegistry`), the same value the markers are drawn from and the one written into
 * numbering.xml, so a list goes out as what it was started as.
 */

import { elementXml, type XmlAttr } from "../ooxml/element";
import { wName } from "../ooxml/names";
import { orderedElement } from "../ooxml/props";
import type {
  LevelIndent,
  LevelSuffix,
  NewList,
  NewListLevel,
  NumberFormat,
  Numbering,
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

function templateLevel(kind: ListKind, ilvl: number): NewListLevel {
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
    run: null,
  };
}

/** The definition a list of this kind is started with */
export function templateList(kind: ListKind): NewList {
  const levels = new Map<number, NewListLevel>();
  for (let ilvl = 0; ilvl < LEVEL_COUNT; ilvl += 1) {
    levels.set(ilvl, templateLevel(kind, ilvl));
  }
  return { levels };
}

/**
 * The definition standing behind this number: the one the document wrote down, or the one a list
 * started while editing was registered with. Undefined for a number nothing defines, which is a
 * list neither drawn nor written.
 */
export function listFor(
  numbering: Numbering,
  numId: number
): NumberingList | undefined {
  return numbering.lists.get(numId) ?? numbering.added.get(numId);
}

function highest(groups: readonly Iterable<number>[]): number {
  let max = 0;
  for (const group of groups) {
    for (const numId of group) if (numId > max) max = numId;
  }
  return max;
}

/**
 * A number for a new list, and the numbering that now defines it under that number.
 *
 * `used` names the numbers the document already spends, the ones no definition stands behind
 * included: a number a paragraph already wears would otherwise join that paragraph to the new list.
 */
export function allocateList(
  numbering: Numbering,
  used: Iterable<number>,
  list: NewList
): { numId: number; numbering: Numbering } {
  const numId =
    highest([numbering.lists.keys(), numbering.added.keys(), used]) + 1;
  return {
    numId,
    numbering: {
      ...numbering,
      added: new Map([...numbering.added, [numId, list]]),
    },
  };
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

/** One child, written only where the level asks for something other than what OOXML already gives */
type OptionalChild = { name: string; xml: string } | null;

/**
 * `w:lvlRestart` counts levels from one, so the level a restart names goes back out as its number
 * plus one, and the level that never restarts as the 0 of §17.9.10. A level restarted by the one
 * above it, which is what a level saying nothing gets, writes nothing.
 */
function lvlRestartXml(restartAfterLevel: number | null): OptionalChild {
  if (restartAfterLevel === null) return null;
  return {
    name: "lvlRestart",
    xml: elementXml(wName("lvlRestart"), [
      [wName("val"), `${restartAfterLevel + 1}`],
    ]),
  };
}

function isLglXml(legal: boolean): OptionalChild {
  return legal ? { name: "isLgl", xml: elementXml(wName("isLgl"), []) } : null;
}

/** A tab is what §17.9.28 puts between the number and the text of a level that says nothing */
function suffXml(suffix: LevelSuffix): OptionalChild {
  if (suffix === "tab") return null;
  return {
    name: "suff",
    xml: elementXml(wName("suff"), [[wName("val"), suffix]]),
  };
}

/**
 * Writes a level out in full.
 *
 * Everything a `NewListLevel` can hold is written here, which is what lets a registered definition
 * be serialized rather than approximated: a value this could not spell would be a value the file
 * silently lost.
 */
function levelXml(ilvl: number, level: NewListLevel): string {
  const ind = indXml(level.indent);
  const optional = [
    lvlRestartXml(level.restartAfterLevel),
    isLglXml(level.legal),
    suffXml(level.suffix),
    ind === "" ? null : { name: "pPr", xml: ind },
  ];
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
        xml: elementXml(wName("lvlJc"), [[wName("val"), level.align]]),
      },
      ...optional.filter((child) => child !== null),
    ]
  );
}

/** Writes one registered definition out as a `<w:abstractNum>` */
export function abstractNumXml(abstractNumId: number, list: NewList): string {
  const body = Array.from(list.levels.entries())
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
