/**
 * The register of list definitions a document node carries.
 *
 * A list started while editing is registered with the definition it was started with, and the
 * register rides on the document node itself: undo takes a registration back with the edit that
 * made it, and the export writes out what was registered rather than working a definition out for
 * itself. Only what `abstractNumXml` writes in full may be registered, and the value is read back
 * through this file so that nothing else can end up in the register.
 */

import {
  LEVEL_ALIGNS,
  LEVEL_SUFFIXES,
  type LevelIndent,
  type NewList,
  type NewListLevel,
} from "./parseNumbering";
import { isNumberFormat } from "./spellers";

/** The lists registered on one document, by the number each was given */
export type NewLists = ReadonlyMap<number, NewList>;

export const NO_NEW_LISTS: NewLists = new Map();

/** The document attr the register travels on */
export const NEW_LISTS_ATTR = "newLists";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function toIndentSlot(value: unknown): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function toLevelIndent(value: unknown): LevelIndent | null {
  if (!isRecord(value)) return null;
  return {
    startTwips: toIndentSlot(value.startTwips),
    endTwips: toIndentSlot(value.endTwips),
    hangingTwips: toIndentSlot(value.hangingTwips),
    firstLineTwips: toIndentSlot(value.firstLineTwips),
  };
}

/** One level of a registered list, as the document node carries it */
function levelValue(
  ilvl: number,
  level: NewListLevel
): Record<string, unknown> {
  return {
    ilvl,
    format: level.format,
    text: level.text,
    start: level.start,
    indent: level.indent,
    restartAfterLevel: level.restartAfterLevel,
    legal: level.legal,
    suffix: level.suffix,
    align: level.align,
  };
}

function toLevel(value: unknown): { ilvl: number; level: NewListLevel } | null {
  if (!isRecord(value)) return null;
  const { ilvl, format, text, start } = value;
  if (typeof ilvl !== "number" || !Number.isInteger(ilvl)) return null;
  if (typeof format !== "string" || !isNumberFormat(format)) return null;
  if (typeof text !== "string") return null;
  if (typeof start !== "number" || !Number.isInteger(start)) return null;
  const suffix = LEVEL_SUFFIXES.find((known) => known === value.suffix);
  const align = LEVEL_ALIGNS.find((known) => known === value.align);
  if (suffix === undefined || align === undefined) return null;
  const restart = value.restartAfterLevel;
  const restartAfterLevel =
    typeof restart === "number" && Number.isInteger(restart) ? restart : null;
  return {
    ilvl,
    level: {
      format,
      text,
      start,
      indent: toLevelIndent(value.indent),
      restartAfterLevel,
      legal: value.legal === true,
      suffix,
      align,
      run: null,
    },
  };
}

function toList(value: unknown): { numId: number; list: NewList } | null {
  if (!isRecord(value)) return null;
  const { numId, levels } = value;
  if (typeof numId !== "number" || !Number.isInteger(numId)) return null;
  if (!Array.isArray(levels)) return null;
  const read = new Map<number, NewListLevel>();
  for (const entry of levels) {
    const level = toLevel(entry);
    if (level === null) return null;
    read.set(level.ilvl, level.level);
  }
  return read.size === 0 ? null : { numId, list: { levels: read } };
}

/**
 * What the document node carries for a register.
 *
 * The entries run in the order of the numbers they define so that the same register is always the
 * same value, whichever order the lists were started in, and an empty register is the null a
 * document that never started a list carries: a list left is then a list never started.
 */
export function newListsValue(lists: NewLists): unknown {
  if (lists.size === 0) return null;
  return Array.from(lists)
    .sort(([left], [right]) => left - right)
    .map(([numId, list]) => ({
      numId,
      levels: Array.from(list.levels)
        .sort(([left], [right]) => left - right)
        .map(([ilvl, level]) => levelValue(ilvl, level)),
    }));
}

/**
 * The register a document node carries. Empty for a node carrying none, and for one carrying an
 * entry that is not a list definition: a definition the writer cannot read back in full is a
 * definition it must not write half of.
 */
export function newListsOf(value: unknown): NewLists {
  if (!Array.isArray(value)) return NO_NEW_LISTS;
  const lists = new Map<number, NewList>();
  for (const entry of value) {
    const read = toList(entry);
    if (read === null) return NO_NEW_LISTS;
    lists.set(read.numId, read.list);
  }
  return lists;
}

/**
 * The register with the definitions of the lists nothing wears any more taken out.
 *
 * A number goes back into circulation with its definition, so leaving a list is exactly what
 * joining it was not: pressing the list button twice gives back the document that was there
 * before the first press.
 */
export function listsWorn(
  lists: NewLists,
  worn: ReadonlySet<number>
): NewLists {
  const dropped = [...lists.keys()].filter((numId) => !worn.has(numId));
  if (dropped.length === 0) return lists;
  return new Map([...lists].filter(([numId]) => worn.has(numId)));
}
