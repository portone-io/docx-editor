/**
 * A composition refused by a locked content control.
 *
 * The lock these tests compose into is put down with the very command the authoring menu runs
 * (`lockSelection`), which writes the same control Word would. The kitchen sink document already carries
 * a control of its own further down, so what the document locked before the command ran is read off
 * first and the new stretch is expected on top of it.
 *
 * The refusal itself is the point and stays: the document may not take the text. What the guard in
 * `src/editor/plugins/lockedContent.ts` adds is the state the editor is left in. The browser drops a
 * composition whose text was taken back out of the DOM under it and sends no `compositionend` for
 * it, so before the guard `view.composing` stayed true for as long as the caret was not moved, and
 * everything the editor holds back for the length of a composition stayed held back.
 */

import { expect, test } from "@playwright/test";
import {
  blocks,
  caretAt,
  composing,
  compositions,
  docText,
  lock,
  lockedText,
  openHarness,
  settle,
} from "./support/harness";
import { commitComposition, compose, imeSession } from "./support/ime";

/** The stretch of the first paragraph the lock goes on, and a spot strictly inside it */
const LOCK_FROM = 2;
const LOCK_LENGTH = 8;
const INSIDE = 5;

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
