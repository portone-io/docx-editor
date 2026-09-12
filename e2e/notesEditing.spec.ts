/**
 * Editing a footnote where it is drawn, at the foot of the page that refers to it.
 *
 * Everything here needs a real layout: the note is drawn in room the page kept for it, which only
 * a browser that laid the text out knows the size of, and the view mounted over it is taken down
 * and put back whenever that room moves to another page. A composition is the case that matters
 * most, since a composition torn down mid-syllable loses what the reader typed.
 */

import { expect, type Page, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  docText,
  noteText,
  openHarness,
  pressModKey,
  rightClick,
  settle,
} from "./support/harness";
import {
  commitComposition,
  compose,
  imeSession,
  setComposition,
} from "./support/ime";

/** The number a footnote is called by, as the body draws it */
function reference(page: Page, label: string) {
  return page.locator(
    `.${editorClassNames.sheet} [aria-label="Footnote ${label}"]`
  );
}

/** The one note an editing view stands over */
function openNote(page: Page) {
  return page.locator(`.${editorClassNames.noteRowOpen}`);
}

/** Waits until the caret is really inside the view over the open note, which is what takes typing */
async function untilNoteHoldsTheCaret(page: Page): Promise<void> {
  await expect(openNote(page).locator(".ProseMirror")).toBeVisible();
  await expect
    .poll(() =>
      page.evaluate(
        (open) => document.activeElement?.closest(`.${open}`) !== null,
        editorClassNames.noteRowOpen
      )
    )
    .toBe(true);
}

async function enterFootnote(page: Page, label: string): Promise<void> {
  await reference(page, label).click();
  await untilNoteHoldsTheCaret(page);
}

test("types, formats, and undoes inside a footnote at the foot of its page", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await enterFootnote(page, "1");

  await page.keyboard.type(" Edited.");
  await expect.poll(() => noteText(page, "1")).toContain("Edited.");
  // The body is untouched by what went into the note
  expect(await docText(page)).not.toContain("Edited.");

  await pressModKey(page, "a");
  await pressModKey(page, "b");
  await expect(
    openNote(page).locator(`.${editorClassNames.run}`).last()
  ).toHaveCSS("font-weight", "700");

  // One history for both surfaces: the undo key inside the note runs the document's own
  await pressModKey(page, "z");
  await expect(
    openNote(page).locator(`.${editorClassNames.run}`).last()
  ).not.toHaveCSS("font-weight", "700");
  await pressModKey(page, "z");
  await expect.poll(() => noteText(page, "1")).not.toContain("Edited.");
});

test("composes hangul inside a footnote while the page lays out again", async ({
  page,
}) => {
  await openHarness(page, "notes");
  // Footnote 2 stands several pages in, so its row is the one a relayout could move
  await enterFootnote(page, "2");
  const cdp = await imeSession(page);

  // 안녕: one composition per syllable, as a 2-beolsik IME delivers it
  await compose(cdp, ["ㅇ", "아", "안"]);
  await settle(page);
  // The page has been measured again in the meantime, and the composition is still open in the
  // very view it was opened in
  expect(
    await page.evaluate(() => window.getSelection()?.rangeCount ?? 0)
  ).toBeGreaterThan(0);
  await expect(openNote(page).locator(".ProseMirror")).toBeVisible();
  await commitComposition(cdp, "안");
  await setComposition(cdp, "ㄴ");
  await compose(cdp, ["녀", "녕"]);
  await commitComposition(cdp, "녕");
  await settle(page);

  await expect.poll(() => noteText(page, "2")).toContain("안녕");
  await expect(openNote(page)).toContainText("안녕");
});

test("returns the caret after the reference on Escape", async ({ page }) => {
  await openHarness(page, "notes");
  await enterFootnote(page, "1");

  await page.keyboard.press("Escape");

  await expect(openNote(page)).toHaveCount(0);
  await expect
    .poll(() =>
      page.evaluate(
        (sheet) => document.activeElement?.classList.contains(sheet) === true,
        editorClassNames.sheet
      )
    )
    .toBe(true);
  // Typing lands just past the number rather than anywhere else on the page. The number itself
  // is a node with no text, which the document reads back as a line of its own
  await page.keyboard.type("!");
  await expect
    .poll(() => docText(page))
    .toContain(
      "Paragraph 5 keeps the text running down the page.\n!\nParagraph 6"
    );
});

test("inserts a footnote from the right click menu and puts the caret in it", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await page.locator(`.${editorClassNames.sheet} p`).first().click();

  await rightClick(page);
  await page.getByRole("menuitem", { name: "Insert footnote" }).click();

  await untilNoteHoldsTheCaret(page);
  await page.keyboard.type("A note written from the menu.");
  await expect
    .poll(() => noteText(page, "3"))
    .toContain("A note written from the menu.");
  await expect(
    page.getByRole("region", { name: "Footnotes on page 1" })
  ).toContainText("A note written from the menu.");
});

test("inserts a footnote on Mod+Alt+F and puts the caret in it", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await page.locator(`.${editorClassNames.sheet} p`).first().click();

  const modifier = process.platform === "darwin" ? "Meta" : "Control";
  await page.keyboard.press(`${modifier}+Alt+f`);

  await untilNoteHoldsTheCaret(page);
  await page.keyboard.type("A note written where it was added.");
  await expect
    .poll(() => noteText(page, "3"))
    .toContain("A note written where it was added.");
  // It is the first footnote of the document now, so the two the file arrived with move up
  await expect(reference(page, "1")).toHaveCount(1);
  // It went in on the first page, so it is drawn at the foot of that page
  await expect(
    page.getByRole("region", { name: "Footnotes on page 1" })
  ).toContainText("A note written where it was added.");
});
