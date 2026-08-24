/**
 * Korean hangul composition, driven as a 2-beolsik IME delivers it.
 *
 * Hangul is composed rather than converted, and that is what sets it apart from the scripts covered
 * in `./composition.spec.ts`. A kana or a pinyin buffer grows to the end of a word and is then
 * swapped wholesale for the candidate the operating system chose. A hangul buffer is one syllable
 * being built out of the jamo of the keys struck: 안 is ㅇ, then ㅏ making 아, then ㄴ closing it.
 * Three things follow from that which no covered script exercises.
 *
 * A syllable settles on its own. The buffer holds one syllable at a time, so the moment a keystroke
 * belongs to the next syllable the IME commits the one standing and opens a fresh composition. A
 * word is therefore as many compositions as it has syllables rather than one, and the editor has to
 * take a run of them back to back without losing a syllable between two of them.
 *
 * A jamo can leave the syllable it went into. Type ㄱㅏㄴ and the buffer reads 간; type ㅏ after it
 * and the ㄴ is no longer the batchim of 간 but the lead of 나, so what stood as one syllable becomes
 * two. Both stagings this reaches the page as are driven below, because the editor may not care
 * which one it is handed: the IME settling 가 and opening a composition for 나, which is what the
 * platform IMEs on macOS and Windows do, and the whole buffer being replaced by 가나 inside one
 * composition, which is all the web composition model requires of an IME.
 *
 * Backspace deletes a jamo rather than a character. It regresses the buffer - 안 to 아 to ㅇ to
 * nothing - and the browser learns of it through an `imeSetComposition` carrying less text each
 * time, not through a keydown the editor could answer.
 */

import { expect, test } from "@playwright/test";
import {
  blocks,
  type CompositionCounts,
  caretAt,
  composing,
  compositions,
  docText,
  firstTextParagraph,
  lock,
  lockedText,
  openHarness,
  settle,
} from "./support/harness";
import {
  clearComposition,
  commitComposition,
  compose,
  imeSession,
  setComposition,
} from "./support/ime";

/** One syllable of a word: the buffers it passes through, and what the IME settles it as */
interface Syllable {
  /** The buffer as it stands after each keystroke, which is one jamo at a time */
  stages: readonly string[];
  /** What the commit lands when the syllable is done */
  settled: string;
}

interface Word {
  syllables: readonly Syllable[];
  /** The text the whole word leaves behind */
  text: string;
}

/** 안녕, one composition per syllable: ㅇ 아 안, settled, then ㄴ 녀 녕 */
const ANNYEONG: Word = {
  syllables: [
    { stages: ["ㅇ", "아", "안"], settled: "안" },
    { stages: ["ㄴ", "녀", "녕"], settled: "녕" },
  ],
  text: "안녕",
};

/**
 * 가나 with the ㄴ moving forward, staged the way macOS and Windows both stage it: the ㅏ that takes
 * the ㄴ off 간 settles 가 and opens a composition holding 나 outright, so ㄴ never stands alone as a
 * buffer of the second syllable. What makes this the interesting one is that the commit lands less
 * text (가) than the buffer it replaces (간).
 */
const GANA_SETTLED: Word = {
  syllables: [
    { stages: ["ㄱ", "가", "간"], settled: "가" },
    { stages: ["나"], settled: "나" },
  ],
  text: "가나",
};

/** 가나 with the ㄴ moving forward inside a single composition, 간 replaced by 가나 in one update */
const GANA_IN_ONE_BUFFER: Word = {
  syllables: [{ stages: ["ㄱ", "가", "간", "가나"], settled: "가나" }],
  text: "가나",
};

const WORDS: ReadonlyArray<[string, Word]> = [
  ["settled one syllable at a time", ANNYEONG],
  ["whose batchim moves forward between two compositions", GANA_SETTLED],
  ["whose batchim moves forward inside one composition", GANA_IN_ONE_BUFFER],
];

/** Every buffer the word passes through, each carrying the syllables settled before it */
function drafts(word: Word): string[] {
  const seen: string[] = [];
  let settled = "";
  for (const syllable of word.syllables) {
    for (const stage of syllable.stages) seen.push(settled + stage);
    settled += syllable.settled;
  }
  return seen;
}

/** What may not be left standing anywhere once the word has landed */
function leftovers(word: Word, original: string): string[] {
  return drafts(word)
    .filter((draft) => draft !== word.text)
    .map((draft) => draft + original);
}

/** One `compositionupdate` per buffer, and one more for the commit that closes each syllable */
function expectedCounts(word: Word): CompositionCounts {
  return {
    start: word.syllables.length,
    update: word.syllables.reduce(
      (sum, syllable) => sum + syllable.stages.length + 1,
      0
    ),
    end: word.syllables.length,
  };
}

for (const [name, word] of WORDS) {
  test(`a hangul word ${name} lands exactly once`, async ({ page }) => {
    await openHarness(page, "kitchen-sink");
    const target = firstTextParagraph(await blocks(page));
    const original = target.docText;
    const rest = (await blocks(page))
      .filter((block) => block.index !== target.index)
      .map((block) => block.docText);
    await caretAt(page, target.index, 0);

    const cdp = await imeSession(page);
    let settled = "";
    for (const [index, syllable] of word.syllables.entries()) {
      for (const stage of syllable.stages) {
        await setComposition(cdp, stage);
        const drawn = (await blocks(page))[target.index];
        expect(drawn?.docText, `document at buffer ${stage}`).toBe(
          settled + stage + original
        );
        expect(drawn?.domText, `screen at buffer ${stage}`).toBe(
          settled + stage + original
        );
        expect(await composing(page)).toBe(true);
      }

      await commitComposition(cdp, syllable.settled);
      settled += syllable.settled;
      const drawn = (await blocks(page))[target.index];
      expect(drawn?.docText, `document once ${settled} is settled`).toBe(
        settled + original
      );
      // The syllable that settled took its composition down with it, and no other one went down
      expect(await composing(page)).toBe(false);
      expect((await compositions(page)).end).toBe(index + 1);
    }

    await settle(page);
    const after = await blocks(page);
    expect(after[target.index]?.docText).toBe(`${word.text}${original}`);
    expect(after[target.index]?.domText).toBe(`${word.text}${original}`);
    // As many compositions as syllables, each opened once and closed once
    expect(await compositions(page)).toEqual(expectedCounts(word));

    // Neither a half-built syllable nor the text that stood here is left anywhere
    const text = await docText(page);
    expect(text.split(word.text)).toHaveLength(2);
    for (const leftover of leftovers(word, original)) {
      expect(text).not.toContain(leftover);
    }
    expect(
      after
        .filter((block) => block.index !== target.index)
        .map((block) => block.docText)
    ).toEqual(rest);
  });
}

/** The buffer as ㅇ, ㅏ and ㄴ are struck */
const ASSEMBLING = ["ㅇ", "아", "안"];

/** The same buffer as backspace takes the jamo back off it, one at a time */
const BACKSPACING = ["아", "ㅇ"];

test("a composition backspaced away leaves nothing committed", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const target = firstTextParagraph(await blocks(page));
  const original = target.docText;
  const before = await docText(page);
  await caretAt(page, target.index, 0);

  const cdp = await imeSession(page);
  await compose(cdp, ASSEMBLING);
  expect((await blocks(page))[target.index]?.docText).toBe(`안${original}`);

  for (const buffer of BACKSPACING) {
    await setComposition(cdp, buffer);
    const drawn = (await blocks(page))[target.index];
    expect(drawn?.docText, `document at buffer ${buffer}`).toBe(
      buffer + original
    );
    expect(drawn?.domText, `screen at buffer ${buffer}`).toBe(
      buffer + original
    );
    // The buffer shrinking is the same composition still standing, not a new one
    expect(await composing(page)).toBe(true);
    expect(await compositions(page)).toMatchObject({ start: 1, end: 0 });
  }

  await clearComposition(cdp);
  await settle(page);

  // Nothing was committed: the document reads as it did before a key was struck
  expect(await docText(page)).toBe(before);
  const after = await blocks(page);
  expect(after[target.index]?.docText).toBe(original);
  expect(after[target.index]?.domText).toBe(original);
  // And the emptied buffer ended its composition rather than leaving it open
  expect(await composing(page)).toBe(false);
  expect(await compositions(page)).toEqual({
    start: 1,
    update: ASSEMBLING.length + BACKSPACING.length + 1,
    end: 1,
  });
});

/** The stretch of the first paragraph the lock goes on, and a spot strictly inside it */
const LOCK_FROM = 2;
const LOCK_LENGTH = 8;
const INSIDE = 5;

test("a hangul composition inside a locked control changes nothing", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const original = (await blocks(page))[0]?.docText ?? "";
  const carried = await lockedText(page);

  expect(await lock(page, 0, LOCK_FROM, LOCK_LENGTH)).toBe(true);
  const locked = await lockedText(page);
  // The new control stands in the first block, ahead of the one the document brought with it
  expect(locked).toBe(
    original.slice(LOCK_FROM, LOCK_FROM + LOCK_LENGTH) + carried
  );

  await caretAt(page, 0, INSIDE);
  const cdp = await imeSession(page);
  await compose(cdp, ASSEMBLING);
  await commitComposition(cdp, "안");
  await settle(page);

  // The document never took the syllable, and the control still holds what it held
  expect((await blocks(page))[0]?.docText).toBe(original);
  expect((await blocks(page))[0]?.domText).toBe(original);
  expect(await lockedText(page)).toBe(locked);

  // Not a jamo of the refused buffer is left anywhere either
  const text = await docText(page);
  for (const buffer of [...ASSEMBLING, "안"]) {
    expect(text).not.toContain(buffer);
  }

  // The refused composition is over rather than standing open for want of a compositionend
  expect(await composing(page)).toBe(false);
  expect((await compositions(page)).end).toBe(0);
});
