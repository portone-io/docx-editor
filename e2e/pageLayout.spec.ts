/**
 * The page measurement against an open composition.
 *
 * Every measurement hands the view fresh pagination decorations. Text composed into an early
 * paragraph shifts every block below it, which moves those marks, so a document of more than one
 * page would remeasure over and over while a composition is open. The guard in
 * `src/page/usePageLayout.ts` holds those measurements back until the composition is over.
 *
 * Without it the page layout moves under the composition while the kana are still uncommitted.
 */

import { expect, test } from "@playwright/test";
import {
  blocks,
  caretAt,
  composing,
  compositions,
  openHarness,
  pagination,
  settle,
} from "./support/harness";
import { commitComposition, compose, imeSession } from "./support/ime";

/** Long enough that the paragraph it goes into takes an extra line, which moves every push below */
const LONG =
  "にほんごのぶんしょうをここにながくかきつづけていくとぎょうがおれる";

const COMMITTED = "日本語の文章をここに長く書き続けていくと行が折れる";

/** The buffer as it stands after every few keystrokes of the sentence above */
const STAGES = [1, 2, 4, 8, 16, 24, 32, LONG.length].map((at) =>
  LONG.slice(0, at)
);

test("the pagination stands still while a composition is open", async ({
  page,
}) => {
  // A document of more than one page, so there is a pagination mark to watch at all
  await openHarness(page, "size-fallback");
  const settled = await pagination(page);
  expect(settled, "the fixture has to reach a second page").not.toBe("");

  await caretAt(page, 0, 0);
  const cdp = await imeSession(page);
  await compose(cdp, STAGES);
  // Two frames is where a measurement would have landed
  await settle(page);
  await settle(page);

  expect(await pagination(page)).toBe(settled);
  expect(await composing(page)).toBe(true);
  expect(await compositions(page)).toEqual({
    start: 1,
    update: STAGES.length,
    end: 0,
  });
  expect((await blocks(page))[0]?.docText.startsWith(LONG)).toBe(true);

  // The measurements held back over the composition come to one, taken once it is over
  await commitComposition(cdp, COMMITTED);
  await settle(page);
  await settle(page);

  expect(await composing(page)).toBe(false);
  expect(await pagination(page)).not.toBe(settled);
  expect((await blocks(page))[0]?.docText.startsWith(COMMITTED)).toBe(true);
});
