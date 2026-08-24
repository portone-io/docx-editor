/**
 * Walks the paragraphs in order and computes the number to show for each list paragraph.
 *
 * A pure computation that never touches the document model: the same paragraphs with the
 * same numbering definitions always yield the same numbers.
 */

import type { NumberingRef } from "../model/format";
import { listFor } from "./listTemplate";
import {
  type LevelIndentPt,
  levelIndentPt,
  type NumberFormat,
  type Numbering,
  type NumberingList,
} from "./parseNumbering";

/** The number to draw in front of one paragraph */
export interface ListMarker {
  text: string;
  /**
   * The indent defined by the level this paragraph belongs to. Used when the paragraph
   * records none of its own
   */
  indent: LevelIndentPt;
}

const ROMAN: ReadonlyArray<readonly [number, string]> = [
  [1000, "M"],
  [900, "CM"],
  [500, "D"],
  [400, "CD"],
  [100, "C"],
  [90, "XC"],
  [50, "L"],
  [40, "XL"],
  [10, "X"],
  [9, "IX"],
  [5, "V"],
  [4, "IV"],
  [1, "I"],
];

function toRoman(value: number): string {
  let rest = value;
  let out = "";
  for (const [amount, sign] of ROMAN) {
    while (rest >= amount) {
      out += sign;
      rest -= amount;
    }
  }
  return out;
}

/** As in Word, z is followed by aa, then bb */
function toLetters(value: number): string {
  const index = (value - 1) % 26;
  const repeat = Math.floor((value - 1) / 26) + 1;
  return String.fromCharCode(65 + index).repeat(repeat);
}

/**
 * Cap on a rendered marker's length. `w:lvlText` is drawn verbatim, so without a cap a crafted
 * megabyte-long one would render in front of every list paragraph and rebuild on every keystroke.
 */
const MAX_MARKER_CHARS = 64;

/**
 * The largest value each format spells out. A number past it is drawn as a decimal, which
 * is where a format we do not spell lands as well.
 *
 * A letter marker takes one character per 26 counted and a roman one one M per 1000, so a
 * crafted `w:start` of two billion would otherwise build a marker of tens of millions of
 * characters.
 */
const MAX_SPELLED_VALUE: Record<NumberFormat, number> = {
  decimal: Number.POSITIVE_INFINITY,
  bullet: Number.POSITIVE_INFINITY,
  lowerLetter: 26 * MAX_MARKER_CHARS,
  upperLetter: 26 * MAX_MARKER_CHARS,
  lowerRoman: 1000 * MAX_MARKER_CHARS,
};

function formatNumber(value: number, format: NumberFormat): string {
  if (value < 1 || value > MAX_SPELLED_VALUE[format]) return `${value}`;
  switch (format) {
    case "lowerLetter":
      return toLetters(value).toLowerCase();
    case "upperLetter":
      return toLetters(value);
    case "lowerRoman":
      return toRoman(value).toLowerCase();
    default:
      return `${value}`;
  }
}

/** Replaces each `%n` in something like `%1.%2.` with the current number of the nth level */
function fillLevels(
  text: string,
  list: NumberingList,
  counters: Map<number, number>
): string {
  return text.replaceAll(/%([1-9])/g, (_, digit: string) => {
    const ilvl = Number(digit) - 1;
    const level = list.levels.get(ilvl);
    if (!level) return "";
    return formatNumber(counters.get(ilvl) ?? level.start, level.format);
  });
}

/**
 * Advances the numbering by one step at this paragraph.
 * A level seen for the first time takes its start number, and when a shallower level
 * advances, the deeper levels start counting from the beginning again.
 */
function advance(
  counters: Map<number, number>,
  list: NumberingList,
  ilvl: number,
  start: number
): void {
  const current = counters.get(ilvl);
  counters.set(ilvl, current === undefined ? start : current + 1);
  for (const deeper of [...counters.keys(), ...list.levels.keys()]) {
    if (deeper > ilvl) counters.delete(deeper);
  }
}

/**
 * A slot is null for a paragraph that is not a list item, or whose level shape could not
 * be found.
 * A number the document does not know (a list started fresh while editing) is drawn with
 * the standard template's shape.
 */
export function computeMarkers(
  paragraphs: readonly (NumberingRef | null)[],
  numbering: Numbering
): (ListMarker | null)[] {
  const countersByList = new Map<number, Map<number, number>>();

  return paragraphs.map((ref) => {
    if (!ref) return null;
    const list = listFor(numbering, ref.numId);
    const level = list.levels.get(ref.ilvl);
    if (!level) return null;

    let counters = countersByList.get(ref.numId);
    if (!counters) {
      counters = new Map<number, number>();
      countersByList.set(ref.numId, counters);
    }
    advance(counters, list, ref.ilvl, level.start);

    const shape =
      level.format === "bullet"
        ? level.text
        : fillLevels(level.text, list, counters);
    const text = shape.slice(0, MAX_MARKER_CHARS);
    return text ? { text, indent: levelIndentPt(level.indent) } : null;
  });
}
