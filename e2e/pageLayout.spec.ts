/**
 * The page measurement against an open composition, and against a document written in two
 * sections on two different papers.
 *
 * Every measurement hands the view fresh pagination decorations. Text composed into an early
 * paragraph shifts every block below it, which moves those marks, so a document of more than one
 * page would remeasure over and over while a composition is open. The guard in
 * `src/page/usePageLayout.ts` holds those measurements back until the composition is over.
 *
 * Without it the page layout moves under the composition while the kana are still uncommitted.
 */

import { expect, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  type FixturePaper,
  LANDSCAPE_PAPER,
  PORTRAIT_PAPER,
} from "./harness/twoSectionsFixture";
import {
  blocks,
  caretAt,
  composing,
  compositions,
  openHarness,
  pageBands,
  pagination,
  settle,
  sheetHeight,
  sheetWidth,
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

/** CSS draws an inch as 96 pixels and a twip is a 1440th of one */
const PX_PER_TWIP = 96 / 1440;

/** The body one page of this paper leaves, in the pixels the sheet is drawn with */
function bodyHeight(paper: FixturePaper): number {
  return (paper.height - paper.top - paper.bottom) * PX_PER_TWIP;
}

/** The body across the paper, which is the width a table standing on it is fitted to */
function bodyWidth(paper: FixturePaper): number {
  return (paper.width - paper.left - paper.right) * PX_PER_TWIP;
}

/** From the end of one page's body to the top of the next: two margins and the grey gap */
function step(paper: FixturePaper): number {
  return (paper.top + paper.bottom) * PX_PER_TWIP + 24;
}

/** One whole page of this paper as the sheet draws it, the gap opened under it included */
function pageHeight(paper: FixturePaper): number {
  return bodyHeight(paper) + step(paper);
}

const PORTRAIT_WIDTH = PORTRAIT_PAPER.width * PX_PER_TWIP;
/** The band tops carry the measured padding of the sheet, so it is the distances that are read */
const TOLERANCE = 1.5;

test("each section is paginated on the paper it names", async ({ page }) => {
  await openHarness(page, "two-sections");
  const bands = await pageBands(page);

  // A portrait page, then landscape pages: the section break opens the second page and the
  // landscape paper holds it
  expect(
    bands.length,
    "the fixture has to reach a fourth page"
  ).toBeGreaterThan(2);
  expect(bands[1] - bands[0]).toBeCloseTo(pageHeight(LANDSCAPE_PAPER), 0);
  expect(bands[2] - bands[1]).toBeCloseTo(pageHeight(LANDSCAPE_PAPER), 0);
  // Where the paper of the first section had been used for the whole document, the page after
  // the section break would have stood a portrait page tall instead
  expect(
    Math.abs(bands[1] - bands[0] - pageHeight(PORTRAIT_PAPER))
  ).toBeGreaterThan(TOLERANCE);

  // The sheet is padded from the first section's paper, so the deeper margin the landscape
  // section ends on is room the last page only has if the layout stretched the sheet to it
  const lastBand = bands[bands.length - 1] ?? 0;
  expect((await sheetHeight(page)) - lastBand).toBeCloseTo(
    pageHeight(LANDSCAPE_PAPER),
    0
  );

  // One sheet is drawn at one width, which is the first section's: the landscape section is
  // paginated on its own height but drawn on the paper the sheet holds
  expect(await sheetWidth(page)).toBeCloseTo(PORTRAIT_WIDTH, 0);
});

test("a column drag in a landscape section is held to that section's body", async ({
  page,
}) => {
  await openHarness(page, "two-sections");
  const table = page.locator(`.${editorClassNames.table}`).first();
  await table.scrollIntoViewIfNeeded();
  const cell = table.getByRole("cell").last();
  const edge = await cell.boundingBox();
  if (!edge) throw new Error("the table was not drawn");

  // Past the body of the landscape section's own paper, so what stops the drag is that limit
  // rather than where the pointer came to rest
  await page.mouse.move(edge.x + edge.width - 1, edge.y + edge.height / 2);
  await page.mouse.down();
  await page.mouse.move(
    edge.x + edge.width + bodyWidth(LANDSCAPE_PAPER),
    edge.y + edge.height / 2
  );
  await page.mouse.up();

  // The paper is drawn at the first section's width and a table wider than the sheet's body is
  // drawn shrunk to fit it (`styles/editor.css`), so it is the width the document now holds
  const width = await table.evaluate((drawn) =>
    drawn instanceof HTMLElement ? Number.parseFloat(drawn.style.width) || 0 : 0
  );
  expect(width).toBeGreaterThan(bodyWidth(PORTRAIT_PAPER) + TOLERANCE);
  // The drag asked for more than the landscape body, and the limit held it to exactly that
  expect(width).toBeCloseTo(bodyWidth(LANDSCAPE_PAPER), 0);
});
