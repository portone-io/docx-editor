/**
 * Walks the paragraphs in order and computes the number to show for each list paragraph.
 *
 * A pure computation that never touches the document model: the same paragraphs with the
 * same numbering definitions always yield the same numbers.
 */

import type { NumberingRef, RunFormat } from "../model/format";
import { listFor } from "./listTemplate";
import {
  type LevelAlign,
  type LevelIndentPt,
  type LevelSuffix,
  levelIndentPt,
  type Numbering,
  type NumberingLevel,
  type NumberingList,
} from "./parseNumbering";
import { spellNumber } from "./spellers";

/** The number to draw in front of one paragraph */
export interface ListMarker {
  /**
   * The number as it is drawn, the space a level asking for one (`w:suff`) included: a separator
   * made of a character belongs to the marker, while the tab every other level asks for is the
   * width the number sits in rather than anything to draw
   */
  text: string;
  /**
   * The indent defined by the level this paragraph belongs to. Used when the paragraph
   * records none of its own
   */
  indent: LevelIndentPt;
  /** Whether the number sits in a width of its own, which is what a tab suffix asks for */
  suffix: LevelSuffix;
  align: LevelAlign;
  /** The formatting the level puts on the number, which dresses the number and nothing else */
  run: RunFormat | null;
}

/**
 * Cap on a rendered marker's length. `w:lvlText` is drawn verbatim, so without a cap a crafted
 * megabyte-long one would render in front of every list paragraph and rebuild on every keystroke.
 * It bounds the spelled numbers as well, so that a crafted `w:start` cannot grow one either.
 */
const MAX_MARKER_CHARS = 64;

/**
 * Replaces each `%n` in something like `%1.%2.` with the current number of the nth level.
 * A legal level spells every one of them as a decimal, whatever format that level counts in.
 */
function fillLevels(
  text: string,
  list: NumberingList,
  counters: Map<number, number>,
  legal: boolean
): string {
  return text.replaceAll(/%([1-9])/g, (_, digit: string) => {
    const ilvl = Number(digit) - 1;
    const level = list.levels.get(ilvl);
    if (!level) return "";
    const count = counters.get(ilvl) ?? level.start;
    return spellNumber(
      count,
      legal ? "decimal" : level.format,
      MAX_MARKER_CHARS
    );
  });
}

/**
 * The level whose advance sends this one back to its start number, which for a level naming none
 * is the level right above it (§17.9.10).
 *
 * A level naming one no shallower than itself needs no rule of its own: only a shallower level
 * ever sets a restart off, and every one of those is shallower than the level it named too, so it
 * restarts exactly as the default does.
 */
function restartedBy(level: NumberingLevel | undefined, ilvl: number): number {
  return level?.restartAfterLevel ?? ilvl - 1;
}

/**
 * Advances the numbering by one step at this paragraph.
 * A level seen for the first time takes its start number, and when a shallower level
 * advances, the deeper levels that answer to it start counting from the beginning again.
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
    if (deeper > ilvl && restartedBy(list.levels.get(deeper), deeper) >= ilvl) {
      counters.delete(deeper);
    }
  }
}

const MARKER_SEGMENTER = new Intl.Segmenter();

/**
 * The cap applied without cutting a character in half: a plain code-unit slice can split a
 * surrogate pair or a combining sequence and render U+FFFD.
 * The shape may be megabytes long, so only a prefix that can possibly survive the cap is
 * segmented, and a cluster that would carry the marker past the cap is dropped whole.
 */
function capped(shape: string): string {
  if (shape.length <= MAX_MARKER_CHARS) return shape;

  let text = "";
  for (const { segment } of MARKER_SEGMENTER.segment(
    shape.slice(0, MAX_MARKER_CHARS * 2)
  )) {
    if (text.length + segment.length > MAX_MARKER_CHARS) break;
    text += segment;
  }
  return text;
}

/**
 * A slot is null for a paragraph that is not a list item, or whose level shape could not be found.
 * A number nothing defines - neither numbering.xml nor a list started while editing - draws no
 * marker, which is what Word does with a `w:numId` no `w:num` answers.
 */
export function computeMarkers(
  paragraphs: readonly (NumberingRef | null)[],
  numbering: Numbering
): (ListMarker | null)[] {
  const countersByList = new Map<number, Map<number, number>>();

  return paragraphs.map((ref) => {
    if (!ref) return null;
    const list = listFor(numbering, ref.numId);
    const level = list?.levels.get(ref.ilvl);
    if (!list || !level) return null;

    let counters = countersByList.get(ref.numId);
    if (!counters) {
      counters = new Map<number, number>();
      countersByList.set(ref.numId, counters);
    }
    advance(counters, list, ref.ilvl, level.start);

    const shape =
      level.format === "bullet"
        ? level.text
        : fillLevels(level.text, list, counters, level.legal);
    const drawn = capped(shape);
    if (!drawn) return null;
    return {
      text: level.suffix === "space" ? `${drawn} ` : drawn,
      indent: levelIndentPt(level.indent),
      suffix: level.suffix,
      align: level.align,
      run: level.run,
    };
  });
}
