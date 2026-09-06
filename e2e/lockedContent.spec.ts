/**
 * What a guard's refusal leaves behind in a real browser, over a locked content control and over
 * the markers a bookmark range is preserved as.
 *
 * The lock these tests compose into is put down with the very command the authoring menu runs
 * (`lockSelection`), which writes the same control Word would. The kitchen sink document already carries
 * a control of its own further down, so what the document locked before the command ran is read off
 * first and the new stretch is expected on top of it. The bookmark range is one the fixture was
 * written with, since a marker is preserved rather than authored and no command puts one down.
 *
 * The refusal itself is the point and stays: the document may not take the text. What the guard in
 * `src/editor/plugins/lockedContent.ts` adds is the state the editor is left in. The browser drops a
 * composition whose text was taken back out of the DOM under it and sends no `compositionend` for
 * it, so before the guard `view.composing` stayed true for as long as the caret was not moved, and
 * everything the editor holds back for the length of a composition stayed held back.
 *
 * A marker is refused from the same list (`src/schema/guards.ts`), but it is reached by a step a
 * command builds rather than by a composition: a marker draws nothing, so the browser rewrites the
 * text around it and leaves it standing, and what the editor reads back out of the DOM never
 * covers it. So the marker is put to a composition to show that typing in a bookmarked paragraph
 * is not refused, and to a keystroke that does build a step across it to show that the deletion is.
 */

import { expect, test } from "@playwright/test";
import {
  type BlockReport,
  blocks,
  caretAt,
  composing,
  compositions,
  docText,
  lock,
  lockedText,
  openHarness,
  selectText,
  settle,
} from "./support/harness";
import { commitComposition, compose, imeSession } from "./support/ime";

/** The stretch of the first paragraph the lock goes on, and a spot strictly inside it */
const LOCK_FROM = 2;
const LOCK_LENGTH = 8;
const INSIDE = 5;

/** The heading the fixture anchors a bookmark range inside */
const BOOKMARKED_HEADING = "Open questions";

/**
 * A bookmark marker draws nothing on screen and takes one position of its own, which the document
 * reads back as a line break. The heading opens with the range's first marker, so a stretch of
 * three covers that marker and the two characters after it.
 */
const MARKER = "\n";
const ACROSS_MARKER = 3;

/** How many of the range's two ends this text still carries */
function markersIn(text: string): number {
  return [...text].filter((character) => character === MARKER).length;
}

function bookmarkedHeading(found: readonly BlockReport[]): BlockReport {
  const heading = found.find((block) =>
    block.docText.startsWith(MARKER + BOOKMARKED_HEADING)
  );
  if (!heading) throw new Error("the fixture holds no bookmarked heading");
  return heading;
}

test("a composition inside a locked control changes nothing", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const before = await blocks(page);
  const original = before[0]?.docText ?? "";
  const carried = await lockedText(page);

  expect(await lock(page, 0, LOCK_FROM, LOCK_LENGTH)).toBe(true);
  const locked = await lockedText(page);
  // The new control stands in the first block, ahead of the one the document brought with it
  expect(locked).toBe(
    original.slice(LOCK_FROM, LOCK_FROM + LOCK_LENGTH) + carried
  );

  await caretAt(page, 0, INSIDE);
  const cdp = await imeSession(page);
  await compose(cdp, ["に", "にほ", "にほん"]);
  await commitComposition(cdp, "日本語");
  await settle(page);

  // The document never took the text, and the control still holds what it held
  expect((await blocks(page))[0]?.docText).toBe(original);
  expect(await docText(page)).not.toContain("日本語");
  expect(await lockedText(page)).toBe(locked);

  // And the screen shows what the document holds, with no character left over
  expect((await blocks(page))[0]?.domText).toBe(original);

  // The refused composition is over rather than standing open for want of a compositionend
  expect(await composing(page)).toBe(false);
  expect((await compositions(page)).end).toBe(0);
});

test("the open edge of a locked control still takes a composition", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const original = (await blocks(page))[0]?.docText ?? "";
  await lock(page, 0, LOCK_FROM, LOCK_LENGTH);
  const locked = await lockedText(page);

  // A control's edges are left open on purpose, so what is typed against one lands outside it
  const edge = LOCK_FROM + LOCK_LENGTH;
  await caretAt(page, 0, edge);
  const cdp = await imeSession(page);
  await compose(cdp, ["に", "にほ"]);
  await commitComposition(cdp, "日本");
  await settle(page);

  const expected = `${original.slice(0, edge)}日本${original.slice(edge)}`;
  expect((await blocks(page))[0]?.docText).toBe(expected);
  expect((await blocks(page))[0]?.domText).toBe(expected);
  expect(await lockedText(page)).toBe(locked);
  expect(await composing(page)).toBe(false);
});

test("a bookmark marker stands through a composition over it", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const heading = bookmarkedHeading(await blocks(page));
  await selectText(page, heading.index, 0, ACROSS_MARKER);

  const cdp = await imeSession(page);
  await compose(cdp, ["に", "にほ"]);
  await commitComposition(cdp, "日本");
  await settle(page);

  // The browser rewrites the text it can see and steps over the marker, which draws nothing, so
  // the composition lands and the guard has nothing to refuse. What matters is that it stays that
  // way: a composition answered as reaching the marker would refuse every word typed into a
  // bookmarked paragraph.
  const after = (await blocks(page))[heading.index]?.docText ?? "";
  expect(markersIn(after)).toBe(markersIn(heading.docText));
  expect(after).toContain("日本");
  expect(await composing(page)).toBe(false);
});

test("a deletion across a bookmark marker is refused", async ({ page }) => {
  await openHarness(page, "kitchen-sink");
  const original = await docText(page);
  const heading = bookmarkedHeading(await blocks(page));
  await selectText(page, heading.index, 0, ACROSS_MARKER);

  // The keymap builds one step across the whole stretch, which does reach the marker
  await page.keyboard.press("Backspace");
  await settle(page);

  const after = (await blocks(page))[heading.index];
  expect(after?.docText).toBe(heading.docText);
  expect(await docText(page)).toBe(original);

  // And the screen shows what the document holds rather than what the browser did to it
  expect(after?.domText).toBe(heading.domText);
});
