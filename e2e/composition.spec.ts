/**
 * A kana-kanji and a pinyin conversion, driven as the browser really delivers them.
 *
 * What an editor sees of an IME is never the keystrokes. The operating system turns those into one
 * `compositionstart`, a `compositionupdate` per keystroke carrying the whole buffer as it stands,
 * and a commit followed by `compositionend`. The candidate list is drawn by the operating system
 * and the page never learns of it; picking a candidate reaches the page as one more update whose
 * buffer has been replaced wholesale. So replaying the buffer of every keystroke, as these tests
 * do, is what typing looks like from the editor's side of the line.
 *
 * ProseMirror reads an open composition into the document as it goes, so the text under the
 * composition is in the document before the commit. What each stage checks is therefore that the
 * paragraph reads as the buffer and the text that was already there and nothing else: no stage left
 * behind, nothing doubled, nothing dropped. The commit then has to leave the converted text there
 * exactly once.
 *
 * The composition updates carry no key events with them. Sending `Input.dispatchKeyEvent`
 * interleaved with `Input.imeSetComposition` over one CDP session stalls for seconds per event in
 * headless Chrome and eventually takes the renderer down with it, and the editor reads nothing off
 * the keys an IME has already swallowed (ProseMirror drops every keydown while `view.composing`),
 * so the buffer states are what is replayed. A key arriving mid-composition is covered on its own
 * in `./keys.spec.ts`, where a single event is all it takes.
 */

import { expect, test } from "@playwright/test";
import {
  blocks,
  caretAt,
  composing,
  compositions,
  docText,
  firstTextParagraph,
  openHarness,
  settle,
} from "./support/harness";
import { commitComposition, imeSession, setComposition } from "./support/ime";

/**
 * The buffer after each keystroke of "nihongo" on a Japanese IME.
 * The trailing n of "hon" only becomes ん once the g that follows it arrives.
 */
const NIHONGO = [
  "n",
  "に",
  "にh",
  "にほ",
  "にほn",
  "にほん",
  "にほんg",
  "にほんご",
];

/** Picking a candidate, then the one after it, then back again */
const NIHONGO_CANDIDATES = ["日本語", "2本後", "日本語"];

/** The buffer after each keystroke of "zhongwen". Pinyin stays in the buffer as it is typed */
const ZHONGWEN = [
  "z",
  "zh",
  "zho",
  "zhon",
  "zhong",
  "zhongw",
  "zhongwe",
  "zhongwen",
];

interface Conversion {
  /** The buffer as it stands after each keystroke */
  stages: readonly string[];
  /** The candidates picked, in the order they were picked */
  candidates: readonly string[];
  /** The text the commit lands */
  committed: string;
  /** What may not be left standing anywhere once the commit has landed */
  leftovers: (original: string) => readonly string[];
}

const CONVERSIONS: ReadonlyArray<[string, Conversion]> = [
  [
    "kana-kanji",
    {
      stages: NIHONGO,
      candidates: NIHONGO_CANDIDATES,
      committed: "日本語",
      leftovers: (original) =>
        [...NIHONGO, "2本後"].map((buffer) => buffer + original),
    },
  ],
  [
    "pinyin",
    {
      stages: ZHONGWEN,
      candidates: ["中文"],
      committed: "中文",
      leftovers: () => ["zhong"],
    },
  ],
];

for (const [name, conversion] of CONVERSIONS) {
  test(`a ${name} conversion lands exactly once`, async ({ page }) => {
    await openHarness(page, "kitchen-sink");
    const target = firstTextParagraph(await blocks(page));
    const original = target.docText;
    const rest = (await blocks(page))
      .filter((block) => block.index !== target.index)
      .map((block) => block.docText);
    await caretAt(page, target.index, 0);

    const cdp = await imeSession(page);
    for (const buffer of [...conversion.stages, ...conversion.candidates]) {
      await setComposition(cdp, buffer);
      const drawn = (await blocks(page))[target.index];
      expect(drawn?.docText, `document at buffer ${buffer}`).toBe(
        buffer + original
      );
      expect(drawn?.domText, `screen at buffer ${buffer}`).toBe(
        buffer + original
      );
      expect(await composing(page)).toBe(true);
      expect((await compositions(page)).end).toBe(0);
    }

    await commitComposition(cdp, conversion.committed);
    await settle(page);

    const after = await blocks(page);
    expect(after[target.index]?.docText).toBe(
      `${conversion.committed}${original}`
    );
    expect(after[target.index]?.domText).toBe(
      `${conversion.committed}${original}`
    );
    // Not one composition was taken down and started again along the way
    expect(await compositions(page)).toEqual({
      start: 1,
      update: conversion.stages.length + conversion.candidates.length + 1,
      end: 1,
    });
    expect(await composing(page)).toBe(false);
    // Neither a stage of the conversion nor the text that stood here is left anywhere
    const text = await docText(page);
    expect(text.split(conversion.committed)).toHaveLength(2);
    for (const leftover of conversion.leftovers(original)) {
      expect(text).not.toContain(leftover);
    }
    expect(
      after
        .filter((block) => block.index !== target.index)
        .map((block) => block.docText)
    ).toEqual(rest);
  });
}

test("plain characters go in with no composition at all", async ({ page }) => {
  await openHarness(page, "kitchen-sink");
  const target = firstTextParagraph(await blocks(page));
  const original = target.docText;
  await caretAt(page, target.index, 0);

  const cdp = await imeSession(page);
  await commitComposition(cdp, "Plain");
  await settle(page);

  const after = await blocks(page);
  expect(after[target.index]?.docText).toBe(`Plain${original}`);
  expect(after[target.index]?.domText).toBe(`Plain${original}`);
  expect(await compositions(page)).toEqual({ start: 0, update: 0, end: 0 });
});
