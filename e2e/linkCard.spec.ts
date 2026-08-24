/**
 * The card that says where a link points, in a real browser.
 *
 * Nothing in jsdom lays out a line, so `coordsAtPos` there answers nothing and the card falls back
 * on the top of the paper: where it really hangs can only be read here. What it shows, what it
 * offers and when it stands down are pinned in jsdom by `src/ui/LinkCard.test.tsx`; what this adds
 * is that the box lands under the link's own first line and inside the screen, and that its actions
 * participate in the browser's tab order without losing the route back to the paper.
 */

import { expect, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  type BlockReport,
  blocks,
  caretAt,
  openHarness,
  pressModKey,
  settle,
} from "./support/harness";

const ADDRESS = "https://example.com/terms";

/** How far into the paragraph the link starts, and how much of it the link covers */
const AT = 4;
const LENGTH = 9;

/** A paragraph with room for a link and text on both sides of it, which a heading has not */
function longParagraph(found: readonly BlockReport[]): BlockReport {
  const paragraph = found.find(
    (block) => block.type === "paragraph" && block.docText.length > 40
  );
  if (!paragraph) throw new Error("the fixture holds no paragraph long enough");
  return paragraph;
}

test("the card hangs under the first line of the link it is about", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const target = longParagraph(await blocks(page));
  await caretAt(page, target.index, AT);
  for (let step = 0; step < LENGTH; step += 1) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  // The scroll that follows the caret would take an open panel away, so it is waited out first
  await settle(page);

  await pressModKey(page, "k");
  const field = page.locator('[role="dialog"][aria-label="Link"] input');
  await expect(field).toBeFocused();
  await page.keyboard.type(ADDRESS);
  await page.keyboard.press("Enter");
  await settle(page);

  // The selection is still the stretch the link went on, so the card is about that link
  const card = page.locator('[role="group"][aria-label="Link"]');
  await expect(card).toContainText(ADDRESS);

  const cardBox = await card.boundingBox();
  const linkBox = await page
    .locator(`.${editorClassNames.link}`)
    .first()
    .boundingBox();
  const viewport = page.viewportSize();
  if (!cardBox || !linkBox || !viewport) throw new Error("nothing was drawn");

  expect(cardBox.y).toBeGreaterThanOrEqual(linkBox.y + linkBox.height - 1);
  expect(Math.abs(cardBox.x - linkBox.x)).toBeLessThan(2);
  expect(cardBox.x).toBeGreaterThan(0);
  expect(cardBox.x + cardBox.width).toBeLessThan(viewport.width);
});

/** Typing inside the link moves the caret on every keystroke, and the card must not follow it */
test("the card stays where it is while text is typed inside the link", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const target = longParagraph(await blocks(page));
  await caretAt(page, target.index, AT);
  for (let step = 0; step < LENGTH; step += 1) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  await settle(page);
  await pressModKey(page, "k");
  await page.keyboard.type(ADDRESS);
  await page.keyboard.press("Enter");
  await settle(page);

  const card = page.locator('[role="group"][aria-label="Link"]');
  const before = await card.boundingBox();
  await caretAt(page, target.index, AT + 2);
  await page.keyboard.type("xx");
  await settle(page);
  const after = await card.boundingBox();

  expect(after).toEqual(before);
});

test("the card is reached by Tab and returns focus to the paper on Escape", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const target = longParagraph(await blocks(page));
  await caretAt(page, target.index, AT);
  for (let step = 0; step < LENGTH; step += 1) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  await settle(page);
  await pressModKey(page, "k");
  await page.keyboard.type(ADDRESS);
  await page.keyboard.press("Enter");
  await settle(page);

  const card = page.locator('[role="group"][aria-label="Link"]');
  await expect(card).toBeVisible();
  await page.keyboard.press("Tab");
  await expect(card.getByRole("button", { name: "Open" })).toBeFocused();

  await page.keyboard.press("Escape");
  await expect(card).toHaveCount(0);
  expect(
    await page.evaluate(
      (sheet) => document.activeElement?.classList.contains(sheet),
      editorClassNames.sheet
    )
  ).toBe(true);
});

test("the card shrinks without introducing its own scrollbars", async ({
  page,
}) => {
  await page.setViewportSize({ width: 320, height: 240 });
  await openHarness(page, "tabs");
  await caretAt(page, 8, 1);
  await settle(page);

  const card = page.locator('[role="group"][aria-label="Link"]');
  await expect(card).toBeVisible();
  await expect(card).toHaveCSS("overflow", "visible");
  const metrics = await card.evaluate((element) => {
    if (!(element instanceof HTMLElement)) {
      throw new Error("the link card is not an HTML element");
    }
    return {
      clientHeight: element.clientHeight,
      clientWidth: element.clientWidth,
      offsetHeight: element.offsetHeight,
      offsetWidth: element.offsetWidth,
    };
  });
  expect(metrics.offsetWidth - metrics.clientWidth).toBe(2);
  expect(metrics.offsetHeight - metrics.clientHeight).toBe(2);
});
