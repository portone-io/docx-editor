/**
 * Writing over everything an inline content control holds, through the gestures a user selects it
 * with and the ways new text arrives.
 *
 * A template marks a field with a bracketed placeholder inside a control, and filling it in means
 * selecting the placeholder and writing over it. The control is a mark on the placeholder's text,
 * so nothing but the plugin keeping it (`src/editor/plugins/controlContents.ts`) stands between
 * that and the new text landing beside the control while the control goes. Each gesture here
 * reaches the model its own way - a drag, a Shift-click, a triple click, typed keys, an IME, a
 * paste - so each is driven for real rather than set through a transaction.
 */

import { expect, type Page, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import { INLINE_PLACEHOLDER } from "./harness/inlineControlFixture";
import {
  composing,
  docText,
  type InlineControlReport,
  inlineControls,
  openHarness,
  selection,
  settle,
} from "./support/harness";
import { commitComposition, compose, imeSession } from "./support/ime";

/** The controls of the fixture, by their tags, in the order they stand */
const START = 0;
const PERIOD = 1;
const SIGNED = 2;

async function controlNamed(
  page: Page,
  tag: string
): Promise<InlineControlReport> {
  const found = (await inlineControls(page)).find(
    (control) => control.tag === tag
  );
  if (!found) throw new Error(`no control tagged ${tag}`);
  return found;
}

/** Where the text of one control is drawn, as the points a pointer takes it by */
async function edgesOf(
  page: Page,
  index: number
): Promise<{ left: number; right: number; y: number }> {
  const box = await page
    .locator(`span.${editorClassNames.sdt}`)
    .nth(index)
    .boundingBox();
  if (!box) throw new Error(`control ${index} is not drawn`);
  return {
    left: box.x + 1,
    right: box.x + box.width - 1,
    y: box.y + box.height / 2,
  };
}

async function drag(
  page: Page,
  from: { x: number; y: number },
  to: { x: number; y: number }
): Promise<void> {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 8 });
  await page.mouse.up();
  await settle(page);
}

/** Waits until the selection the gesture made has reached the editor, which a drag does late */
async function selected(page: Page): Promise<void> {
  await expect
    .poll(async () => {
      const { from, to } = await selection(page);
      return to > from;
    })
    .toBe(true);
}

test.beforeEach(async ({ page }) => {
  await openHarness(page, "inline-controls");
});

test("a drag over the placeholder, typed over, leaves the text in the control", async ({
  page,
}) => {
  const edges = await edgesOf(page, START);
  await drag(
    page,
    { x: edges.left, y: edges.y },
    { x: edges.right, y: edges.y }
  );
  await selected(page);
  await page.keyboard.type("2026-10-01");
  await settle(page);

  expect(await controlNamed(page, "START")).toEqual({
    tag: "START",
    text: "2026-10-01",
    empty: false,
  });
  expect((await docText(page)).split("\n")[0]).toBe(
    "Starts on 2026-10-01 at noon."
  );
});

test("a drag made backwards reads the same", async ({ page }) => {
  const edges = await edgesOf(page, START);
  await drag(
    page,
    { x: edges.right, y: edges.y },
    { x: edges.left, y: edges.y }
  );
  await selected(page);
  await page.keyboard.type("2026");
  await settle(page);

  expect((await controlNamed(page, "START")).text).toBe("2026");
});

test("a drag carried past the end of the line keeps a control standing alone", async ({
  page,
}) => {
  const edges = await edgesOf(page, PERIOD);
  await drag(
    page,
    { x: edges.left, y: edges.y },
    { x: edges.right + 120, y: edges.y }
  );
  await selected(page);
  await page.keyboard.type("12 months");
  await settle(page);

  expect((await controlNamed(page, "PERIOD")).text).toBe("12 months");
});

test("a Shift-click past the end of the line does the same", async ({
  page,
}) => {
  const edges = await edgesOf(page, PERIOD);
  await page.mouse.click(edges.left, edges.y);
  await page.keyboard.down("Shift");
  await page.mouse.click(edges.right + 120, edges.y);
  await page.keyboard.up("Shift");
  await settle(page);
  await selected(page);
  await page.keyboard.type("12 months");
  await settle(page);

  expect((await controlNamed(page, "PERIOD")).text).toBe("12 months");
});

test("a triple click over a control standing alone does the same", async ({
  page,
}) => {
  const edges = await edgesOf(page, PERIOD);
  await page.mouse.click((edges.left + edges.right) / 2, edges.y, {
    clickCount: 3,
  });
  await settle(page);
  await selected(page);
  await page.keyboard.type("12 months");
  await settle(page);

  expect((await controlNamed(page, "PERIOD")).text).toBe("12 months");
});

test("a hangul word composed over the placeholder lands in the control, syllable by syllable", async ({
  page,
}) => {
  const edges = await edgesOf(page, START);
  await drag(
    page,
    { x: edges.left, y: edges.y },
    { x: edges.right, y: edges.y }
  );
  await selected(page);

  const cdp = await imeSession(page);
  await compose(cdp, ["ㅎ", "하", "한"]);
  // The composition is still open under the caret: keeping the control did not take it down
  expect(await composing(page)).toBe(true);
  await commitComposition(cdp, "한");
  await settle(page);
  await compose(cdp, ["ㄱ", "그", "글"]);
  expect(await composing(page)).toBe(true);
  await commitComposition(cdp, "글");
  await settle(page);

  expect((await controlNamed(page, "START")).text).toBe("한글");
  expect((await docText(page)).split("\n")[0]).toBe("Starts on 한글 at noon.");
});

test("deleting the placeholder leaves the control standing empty, and typing fills it again", async ({
  page,
}) => {
  const edges = await edgesOf(page, START);
  await drag(
    page,
    { x: edges.left, y: edges.y },
    { x: edges.right, y: edges.y }
  );
  await selected(page);
  await page.keyboard.press("Backspace");
  await settle(page);

  expect(await controlNamed(page, "START")).toEqual({
    tag: "START",
    text: "",
    empty: true,
  });

  await page.keyboard.type("Monday");
  await settle(page);
  expect(await controlNamed(page, "START")).toEqual({
    tag: "START",
    text: "Monday",
    empty: false,
  });
});

test("a paste over the placeholder lands in the control", async ({ page }) => {
  const edges = await edgesOf(page, START);
  await drag(
    page,
    { x: edges.left, y: edges.y },
    { x: edges.right, y: edges.y }
  );
  await selected(page);
  await page.evaluate((sheetClass) => {
    const data = new DataTransfer();
    data.setData("text/plain", "2026-10-01");
    document.querySelector(`.${sheetClass}`)?.dispatchEvent(
      new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData: data,
      })
    );
  }, editorClassNames.sheet);
  await settle(page);

  expect((await controlNamed(page, "START")).text).toBe("2026-10-01");
});

test("a control locked against deletion alone takes the text typed over it", async ({
  page,
}) => {
  const edges = await edgesOf(page, SIGNED);
  await drag(
    page,
    { x: edges.left, y: edges.y },
    { x: edges.right, y: edges.y }
  );
  await selected(page);
  await page.keyboard.type("Friday");
  await settle(page);

  expect((await controlNamed(page, "SIGNED")).text).toBe("Friday");
  expect((await docText(page)).split("\n")).toContain("Signed on Friday.");
});

test("the placeholder the tests write over is what the fixture holds", async ({
  page,
}) => {
  expect((await inlineControls(page)).map((control) => control.text)).toEqual([
    INLINE_PLACEHOLDER,
    INLINE_PLACEHOLDER,
    INLINE_PLACEHOLDER,
  ]);
});
