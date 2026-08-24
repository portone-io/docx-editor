/**
 * A key arriving while a composition is open.
 *
 * On a Japanese IME, Enter is what settles the candidate, and the browser sends the keydown for it
 * to the page as well. Were the editor to answer that keydown, `src/editor/plugins/keymap.ts` would split
 * the paragraph underneath the composition. ProseMirror drops every keydown while
 * `view.composing`, and this is where that is held to.
 */

import { expect, test } from "@playwright/test";
import {
  blocks,
  caretAt,
  composing,
  compositions,
  openHarness,
  settle,
} from "./support/harness";
import {
  commitComposition,
  compose,
  ENTER,
  imeSession,
  pressKey,
} from "./support/ime";

test("enter does not split the paragraph under a composition", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const before = await blocks(page);
  const original = before[0]?.docText ?? "";
  await caretAt(page, 0, 0);

  const cdp = await imeSession(page);
  await compose(cdp, ["に", "にほ", "にほん"]);
  await pressKey(cdp, ENTER);
  await settle(page);

  expect(await blocks(page)).toHaveLength(before.length);
  expect(await composing(page)).toBe(true);
  expect((await compositions(page)).end).toBe(0);

  await commitComposition(cdp, "日本語");
  await settle(page);

  const after = await blocks(page);
  expect(after).toHaveLength(before.length);
  expect(after[0]?.docText).toBe(`日本語${original}`);
  expect(after[0]?.domText).toBe(`日本語${original}`);
});
