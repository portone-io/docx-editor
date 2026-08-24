/**
 * A page break splitting the block it stands in, in a real browser.
 *
 * Nothing in jsdom lays out a line, so where a break falls on the page and where the caret is
 * drawn afterwards can only be read here. The arithmetic is pinned by `src/page/pageLayout.test.ts`
 * and the reading of the sheet by `src/page/measureBlocks.test.ts`; what this adds is that the
 * space really does land the text and the caret on the next page.
 */

import { expect, type Page, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  type BlockReport,
  blockHeight,
  blocks,
  caretAt,
  caretBox,
  firstTextParagraph,
  openHarness,
  settle,
  spaces,
} from "./support/harness";

/** How far into the line the break is put, leaving text on either side of it */
const AT = 12;

/** The gap drawn between the first page and the second, in the coordinates the caret is read in */
function firstBand(
  page: Page
): Promise<{ top: number; bottom: number } | null> {
  return page.evaluate((split) => {
    const band = document.querySelector(`.${split}`);
    if (!band) return null;
    const box = band.getBoundingClientRect();
    return { top: box.top, bottom: box.bottom };
  }, editorClassNames.pageSplit);
}

/** A list item with text on both sides of the break, so the list formatting is in the way */
function listItem(found: readonly BlockReport[]): BlockReport {
  const item = found.find(
    (block) => block.marker !== null && block.docText.length > AT + 20
  );
  if (!item) throw new Error("the fixture holds no list item long enough");
  return item;
}

test("a page break carries the rest of a list item to the next page, caret and all", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const before = await blocks(page);
  const target = listItem(before);
  await caretAt(page, target.index, AT);

  await page.keyboard.press("Meta+Enter");
  await settle(page);
  await settle(page);

  const after = await blocks(page);
  // One item still, with its marker and every character of it
  expect(after).toHaveLength(before.length);
  expect(after[target.index]?.marker).toBe(target.marker);
  expect(after[target.index]?.docText).toBe(
    `${target.docText.slice(0, AT)}\n${target.docText.slice(AT)}`
  );

  // One break, given a real height: the space left on the page it fell on
  const opened = await spaces(page);
  expect(opened).toMatch(/^\d+(\.\d+)?\/\d+(\.\d+)?px$/);
  expect(Number.parseFloat(opened)).toBeGreaterThan(0);

  const band = await firstBand(page);
  expect(band).not.toBeNull();
  // The caret sits after the break, so it is on the page the text after the break went to
  expect((await caretBox(page)).top).toBeGreaterThan(band?.bottom ?? 0);
});

test("the space opened at a break stands still once it is applied", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const target = listItem(await blocks(page));
  await caretAt(page, target.index, AT);

  await page.keyboard.press("Meta+Enter");
  await settle(page);
  await settle(page);
  const settled = await spaces(page);

  // Applying the space resizes the sheet, which asks for another measurement: the space it
  // works out has to be the one already there, or the page would creep on every pass
  for (let pass = 0; pass < 3; pass += 1) await settle(page);
  expect(await spaces(page)).toBe(settled);
});

test("taking the break out again takes its space with it", async ({ page }) => {
  await openHarness(page, "kitchen-sink");
  const target = listItem(await blocks(page));
  await caretAt(page, target.index, AT);

  await page.keyboard.press("Meta+Enter");
  await settle(page);
  await settle(page);
  expect(await spaces(page)).not.toBe("");

  // The caret stands right after the break it just put in
  await page.keyboard.press("Backspace");
  await settle(page);
  await settle(page);

  expect(await spaces(page)).toBe("");
  expect((await blocks(page))[target.index]?.docText).toBe(target.docText);
});

/**
 * The idiom Word writes for a break between two paragraphs is a paragraph holding nothing but the
 * break, and an empty space costs that paragraph the line the `br` alone occupied
 * (`src/page/pageDecorations`). What it may not cost it is the whole paragraph: a block drawn at no
 * height is left out of the measurement (`src/page/measureBlocks`), and its break would then never
 * be given a space at all.
 */
test("a paragraph holding nothing but a page break is measured and given its space", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const target = firstTextParagraph(await blocks(page));
  await caretAt(page, target.index, target.docText.length);

  await page.keyboard.press("Enter");
  await page.keyboard.press("Meta+Enter");
  await settle(page);
  await settle(page);

  const standalone = target.index + 1;
  expect((await blocks(page))[standalone]?.docText).toBe("\n");
  expect(await blockHeight(page, standalone)).toBeGreaterThan(0);
  expect(Number.parseFloat(await spaces(page))).toBeGreaterThan(0);
});
