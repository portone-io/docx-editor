/**
 * The blank a form leaves to be filled in: a run with a highlight and no text.
 *
 * It holds no character, so on paper it would be nothing; the editor draws it as a short box in
 * its highlight (`src/styles/editor.css`), and what is written there takes its formatting and
 * replaces it (`src/editor/plugins/controlContents.ts`). Whether the box is wide enough to press,
 * where a press leaves the caret, and whether a hangul composition opened beside an atom that goes
 * away under it lands every syllable once are questions only a real browser answers.
 */

import { expect, type Page, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import { EMPTY_RUN_BLANKS } from "./harness/emptyRunFixture";
import {
  composing,
  docText,
  emptyRuns,
  highlightOf,
  type InlineControlReport,
  inlineControls,
  openHarness,
  settle,
} from "./support/harness";
import {
  commitComposition,
  compose,
  imeSession,
  setComposition,
} from "./support/ime";

const BLANK_COUNT = Object.keys(EMPTY_RUN_BLANKS).length;

const YELLOW = "rgb(255, 255, 0)";

function blank(page: Page, index: number) {
  return page.locator(`.${editorClassNames.emptyRun}`).nth(index);
}

async function controlNamed(
  page: Page,
  tag: string
): Promise<InlineControlReport | undefined> {
  return (await inlineControls(page)).find((control) => control.tag === tag);
}

/** Presses the middle of the left half of a blank, which is the half a caret would land before */
async function pressLeftHalf(page: Page, index: number): Promise<void> {
  const box = await blank(page, index).boundingBox();
  if (!box) throw new Error(`blank ${index} is not drawn`);
  await page.mouse.click(box.x + box.width / 4, box.y + box.height / 2);
  await settle(page);
}

/**
 * The lines of the document, one per block, a table's cells each on a line of their own. A node
 * holding no text, a blank among them, reads as a line break of its own
 */
async function lines(page: Page): Promise<string[]> {
  return (await docText(page)).split("\n");
}

test.beforeEach(async ({ page }) => {
  await openHarness(page, "empty-runs");
});

test("each blank is drawn as a box in its highlight", async ({ page }) => {
  await expect(page.locator(`.${editorClassNames.emptyRun}`)).toHaveCount(
    BLANK_COUNT
  );
  for (let index = 0; index < BLANK_COUNT; index += 1) {
    const drawn = await blank(page, index).evaluate((element) => {
      const box = element.getBoundingClientRect();
      const run = element.closest(".docx-editor-run");
      const runBox = run?.getBoundingClientRect();
      return {
        width: box.width,
        fontSize: Number.parseFloat(getComputedStyle(element).fontSize),
        background: run ? getComputedStyle(run).backgroundColor : null,
        // The run paints the colour, so it has to reach over the whole box
        covered:
          runBox !== undefined &&
          runBox.left <= box.left &&
          runBox.right >= box.right &&
          runBox.height > 0,
      };
    });
    expect(drawn.width).toBeGreaterThan(drawn.fontSize * 0.9);
    expect(drawn.background).toBe(YELLOW);
    expect(drawn.covered).toBe(true);
  }
});

test("a press on a blank in a cell and typing fills it with highlighted text", async ({
  page,
}) => {
  await pressLeftHalf(page, EMPTY_RUN_BLANKS.nameCell);
  await page.keyboard.type("Kim");
  await settle(page);

  expect(await emptyRuns(page)).toBe(BLANK_COUNT - 1);
  expect(await highlightOf(page, "Kim")).toBe("yellow");
  expect(await lines(page)).toContain("Kim");
});

test("a press on the left half of a blank after a label still writes into the blank", async ({
  page,
}) => {
  await pressLeftHalf(page, EMPTY_RUN_BLANKS.label);
  await page.keyboard.type("Lee");
  await settle(page);

  expect(await emptyRuns(page)).toBe(BLANK_COUNT - 1);
  expect(await highlightOf(page, "Lee")).toBe("yellow");
  expect(await lines(page)).toContain("Label: Lee");
});

test("typing into the blank an open control holds writes into the control", async ({
  page,
}) => {
  await pressLeftHalf(page, EMPTY_RUN_BLANKS.signer);
  await page.keyboard.type("Park");
  await settle(page);

  expect(await controlNamed(page, "SIGNER")).toEqual({
    tag: "SIGNER",
    text: "Park",
    empty: false,
  });
  expect(await highlightOf(page, "Park")).toBe("yellow");
});

test("a blank a locked control holds stays, and what is typed stands beside it", async ({
  page,
}) => {
  await pressLeftHalf(page, EMPTY_RUN_BLANKS.shut);
  await page.keyboard.type("X");
  await settle(page);

  expect(await emptyRuns(page)).toBe(BLANK_COUNT);
  expect((await controlNamed(page, "SHUT"))?.text ?? "").toBe("");
  // The blank stays between the label and what was typed after it
  expect(await docText(page)).toContain("Shut: \nX");
});

test("Backspace after a blank takes it away", async ({ page }) => {
  await pressLeftHalf(page, EMPTY_RUN_BLANKS.erase);
  await page.keyboard.press("Backspace");
  await settle(page);

  expect(await emptyRuns(page)).toBe(BLANK_COUNT - 1);
  expect(await lines(page)).toContain("Erase: ");
});

test("a hangul word composed into a blank in a cell lands once, highlighted", async ({
  page,
}) => {
  await pressLeftHalf(page, EMPTY_RUN_BLANKS.roleCell);
  const cdp = await imeSession(page);
  await compose(cdp, ["ㅎ", "하", "한"]);
  // The blank went as the first jamo landed; the composition is still open under the caret
  expect(await composing(page)).toBe(true);
  await commitComposition(cdp, "한");
  await settle(page);
  // A backspace inside the composition regresses the buffer rather than deleting a character
  await compose(cdp, ["ㄱ", "그", "글"]);
  await setComposition(cdp, "그");
  await setComposition(cdp, "글");
  expect(await composing(page)).toBe(true);
  await commitComposition(cdp, "글");
  await settle(page);

  expect(await lines(page)).toContain("한글");
  expect(await highlightOf(page, "한글")).toBe("yellow");
  expect(await emptyRuns(page)).toBe(BLANK_COUNT - 1);
});

test("a hangul word composed into the blank an open control holds stays in the control", async ({
  page,
}) => {
  await pressLeftHalf(page, EMPTY_RUN_BLANKS.witness);
  const cdp = await imeSession(page);
  await compose(cdp, ["ㅎ", "하", "한"]);
  expect(await composing(page)).toBe(true);
  await commitComposition(cdp, "한");
  await settle(page);
  await compose(cdp, ["ㄱ", "그", "글"]);
  await commitComposition(cdp, "글");
  await settle(page);

  expect(await controlNamed(page, "WITNESS")).toEqual({
    tag: "WITNESS",
    text: "한글",
    empty: false,
  });
  expect(await lines(page)).toContain("Witness: 한글");
});
