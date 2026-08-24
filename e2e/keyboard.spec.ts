/**
 * The editor driven by a keyboard alone, in a real browser.
 *
 * Focus and the keys that scroll a page are what jsdom cannot answer for: nothing there decides
 * that a key belongs to a scroll rather than to the widget under it, and nothing draws, so a panel
 * placed against the screen is never really on it. Both are what the interaction layer turns on,
 * so both are driven here.
 *
 * The two things this pins down had no way through at all before.
 * A right click menu inside a table was reachable by mouse alone: Tab belongs to the table, since
 * prosemirror-tables walks the cells with it, and the first arrow key moved the caret, scrolled
 * the paper to follow it, and the scroll took the menu away.
 * The toolbar was thirty stops in the tab order on the way to the paper.
 */

import { expect, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  blocks,
  caretAt,
  caretInCell,
  firstTextParagraph,
  focused,
  openHarness,
  pressModKey,
  rightClick,
  settle,
  tableRows,
} from "./support/harness";

/** The one control the toolbar opens on: nothing has been edited, so undo and redo are dead */
const FIRST_CONTROL = "Zoom";

const scrollY = (page: import("@playwright/test").Page) =>
  page.evaluate(() => window.scrollY);

test("a table menu is reached, walked and run from the keyboard", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  expect(await caretInCell(page)).not.toBe("");
  const rowsBefore = await tableRows(page);

  await rightClick(page);
  const menu = page.locator('[role="menu"][aria-label="Table actions"]');
  await expect(menu).toBeVisible();
  // The menu takes the focus, which is the whole of how a keyboard gets into it
  expect(await focused(page)).toBe("Insert row above");

  await page.keyboard.press("ArrowDown");
  expect(await focused(page)).toBe("Insert row below");
  await page.keyboard.press("Enter");
  await settle(page);

  expect(await tableRows(page)).toBe(rowsBefore + 1);
  await expect(menu).toHaveCount(0);
});

test("a scroll key leaves the open menu where it is", async ({ page }) => {
  await openHarness(page, "kitchen-sink");
  await caretInCell(page);
  await rightClick(page);
  const menu = page.locator('[role="menu"][aria-label="Table actions"]');
  await expect(menu).toBeVisible();
  const restingAt = await scrollY(page);

  for (const key of ["ArrowDown", "ArrowDown", "PageDown", "End"]) {
    await page.keyboard.press(key);
  }

  await expect(menu).toBeVisible();
  // Nothing moved under the menu: the keys belong to it while it stands
  expect(await scrollY(page)).toBe(restingAt);
  // The harness mounts the editor for authoring, so the row that lifts a lock ends the menu
  expect(await focused(page)).toBe("Unlock");
});

test("escape closes the menu and hands the focus back to the paper", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  await caretInCell(page);
  await rightClick(page);
  await expect(
    page.locator('[role="menu"][aria-label="Table actions"]')
  ).toBeVisible();

  await page.keyboard.press("Escape");
  await expect(
    page.locator('[role="menu"][aria-label="Table actions"]')
  ).toHaveCount(0);
  expect(
    await page.evaluate(
      (sheet) => document.activeElement?.classList.contains(sheet),
      editorClassNames.sheet
    )
  ).toBe(true);
});

/**
 * The panel is a form, so an address cannot be typed into a field the keyboard never reached. It is
 * placed before it is shown, and a box still invisible for want of a place takes no focus, which is
 * what jsdom cannot answer for: nothing there draws, so the field takes the focus either way.
 */
test("the link panel takes the focus as it opens, and hands it back on escape", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const target = firstTextParagraph(await blocks(page));
  await caretAt(page, target.index, 0);
  for (let step = 0; step < 4; step += 1) {
    await page.keyboard.press("Shift+ArrowRight");
  }
  // The scroll that follows the caret would take the panel away, so it is waited out first
  await settle(page);

  await pressModKey(page, "k");
  const panel = page.locator('[role="dialog"][aria-label="Link"]');
  await expect(panel).toBeVisible();
  expect(await focused(page)).toBe("Address");

  await page.keyboard.press("Escape");
  await expect(panel).toHaveCount(0);
  expect(
    await page.evaluate(
      (sheet) => document.activeElement?.classList.contains(sheet),
      editorClassNames.sheet
    )
  ).toBe(true);
});

test("the toolbar is one stop on the way to the paper, walked with the arrows", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  await caretAt(page, 0, 0);

  // The toolbar is drawn ahead of the paper, so it is the stop behind it
  await page.keyboard.press("Shift+Tab");
  expect(await focused(page)).toBe(FIRST_CONTROL);

  await page.keyboard.press("ArrowRight");
  expect(await focused(page)).toBe("Style");
  await page.keyboard.press("ArrowRight");
  expect(await focused(page)).toBe("Font");
  await page.keyboard.press("End");
  expect(await focused(page)).toBe("Show comments");
  await page.keyboard.press("Home");
  expect(await focused(page)).toBe(FIRST_CONTROL);

  // One stop for the whole row: the next Tab is out of it and onto the paper
  await page.keyboard.press("Tab");
  expect(
    await page.evaluate(
      (sheet) => document.activeElement?.classList.contains(sheet),
      editorClassNames.sheet
    )
  ).toBe(true);
});

test("a toolbar panel is opened, walked and chosen from with the keyboard", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  await caretAt(page, 0, 0);
  const trigger = page.locator('button[aria-label="Alignment"]');
  await trigger.focus();
  const restingAt = await scrollY(page);

  await page.keyboard.press("Enter");
  const menu = page.locator('[role="menu"][aria-label="Paragraph alignment"]');
  await expect(menu).toBeVisible();
  // The heading the fixture opens with is centered, and the walk starts on the choice in force
  expect(await focused(page)).toBe("Align center");

  await page.keyboard.press("ArrowDown");
  expect(await focused(page)).toBe("Align right");
  expect(await scrollY(page)).toBe(restingAt);

  await page.keyboard.press("Enter");
  await settle(page);
  await expect(menu).toHaveCount(0);
  // The focus comes back to the button the panel hangs from
  expect(await focused(page)).toBe("Alignment");
  expect(
    await page
      .locator(`.${editorClassNames.sheet} p`)
      .first()
      .evaluate((paragraph) => getComputedStyle(paragraph).textAlign)
  ).toBe("right");
});

/**
 * The pointer opening a panel from a button that is already holding the focus.
 *
 * Which of the two opened a panel decides whether it takes the focus and hands it back on the way
 * out, and a click on a toolbar button moves the focus nowhere (`ToolbarButton.tsx`), so where the
 * focus was standing cannot be what tells them apart: the button holds it after an Escape has
 * handed it back, and the pointer open that follows was read as a keyboard one. The panel then took
 * the focus, and closing it pulled the focus off the paper the pick had just handed it back to, so
 * the character typed next was dropped on the toolbar button.
 *
 * A real browser is the only place this shows: a click there moves the focus unless the mousedown
 * is cancelled, which is exactly what the button does.
 */
test("a color picked with the pointer leaves the typing to the paper", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const target = firstTextParagraph(await blocks(page));
  const original = target.docText;
  await caretAt(page, target.index, 0);

  const trigger = page.locator('button[aria-label="Text color"]');
  const palette = page.locator('[role="grid"][aria-label="Colors"]');

  // The keyboard first, so that the button is the one holding the focus by the time the pointer comes
  await trigger.focus();
  await page.keyboard.press("Enter");
  await expect(palette).toBeVisible();
  await page.keyboard.press("Escape");
  await expect(palette).toHaveCount(0);
  expect(await focused(page)).toBe("Text color");

  await trigger.click();
  await expect(palette).toBeVisible();
  // A panel the pointer opened moves no focus at all, so the button is still holding it
  expect(await focused(page)).toBe("Text color");

  await palette.locator('button[aria-label="#ff0000"]').click();
  await expect(palette).toHaveCount(0);
  await settle(page);
  await page.keyboard.type("Z");

  expect((await blocks(page))[target.index]?.docText).toBe(`Z${original}`);
});

test("escape closes a toolbar panel and hands the focus back to its button", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const trigger = page.locator('button[aria-label="Text color"]');
  await trigger.focus();
  await page.keyboard.press("Enter");

  const palette = page.locator('[role="grid"][aria-label="Colors"]');
  await expect(palette).toBeVisible();
  // Nothing is painted, so the walk starts at the entry that withdraws the color
  expect(await focused(page)).toBe("None");
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("ArrowRight");
  expect(await focused(page)).toBe("#434343");

  await page.keyboard.press("Escape");
  await expect(palette).toHaveCount(0);
  expect(await focused(page)).toBe("Text color");
});
