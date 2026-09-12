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
  composing,
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

/** The number a note is called by, as the body draws it */
function reference(page: Page, label: string, kind = "Footnote") {
  return page.locator(
    `.${editorClassNames.sheet} [aria-label="${kind} ${label}"]`
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

/**
 * A composition that writes over the number a note is drawn by.
 *
 * The number goes back into the very paragraph the composition is open in, which is the one edit
 * that puts a node beside composed text rather than rewriting it - the same difference that lets
 * a comment be put back under an open composition (`hangulComposition.spec.ts`). Only a real
 * browser holds a composition, so this is the one place the pairing can be measured.
 */
test("composes over the number a note is drawn by, and keeps it", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await enterFootnote(page, "1");
  const number = openNote(page).locator(`.${editorClassNames.noteMark}`);
  await expect(number).toHaveText("1");
  const cdp = await imeSession(page);

  // Back over every character the note spells, then one step of selection onto the number itself
  const spelled = await noteText(page, "1");
  for (let step = 0; step < spelled.length; step += 1) {
    await page.keyboard.press("ArrowLeft");
  }
  await page.keyboard.press("Shift+ArrowLeft");
  await compose(cdp, ["ㅇ", "아", "안"]);
  await settle(page);
  await commitComposition(cdp, "안");
  await settle(page);

  // What was composed stands where the number was, and the number stands ahead of it
  await expect.poll(() => noteText(page, "1")).toBe(`안${spelled}`);
  await expect(number).toHaveText("1");
  // The composition was let go of rather than left open over a note the restoration moved
  expect(await composing(page)).toBe(false);
});

/**
 * Writing a note again from nothing: select all of it, then type.
 *
 * A composition replaces what is selected because the browser does the replacing, and Chrome will
 * not touch a selection that begins at an element it may not edit. A note always begins with the
 * number it is drawn by, so the whole of a note is exactly such a selection, and the syllables
 * used to land in front of the words they were meant to replace. The selection is taken away
 * before the composition opens now (`editor/plugins/compositionSelection`).
 */
test("writes a note again from nothing with hangul", async ({ page }) => {
  await openHarness(page, "notes");
  await enterFootnote(page, "1");
  const number = openNote(page).locator(`.${editorClassNames.noteMark}`);
  const cdp = await imeSession(page);

  await pressModKey(page, "a");
  await compose(cdp, ["ㅇ", "아", "안"]);
  await settle(page);
  await commitComposition(cdp, "안");
  await settle(page);

  // Nothing of what stood there is left, and the number the note is drawn by still is
  await expect.poll(() => noteText(page, "1")).toBe("안");
  await expect(number).toHaveText("1");
  expect(await composing(page)).toBe(false);
});

/**
 * The number Word draws a note by is a preserved chip inside the note, and a browser deletes an
 * inline atom itself: the deletion arrives as a DOM change read back rather than as a key the
 * keymap answered, which is the path only a real browser takes.
 */
test("keeps the note's own number when its whole text is deleted", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await enterFootnote(page, "1");
  const number = openNote(page).locator(`.${editorClassNames.noteMark}`);
  await expect(number).toHaveText("1");

  await pressModKey(page, "a");
  await page.keyboard.press("Backspace");

  await expect.poll(() => noteText(page, "1")).toBe("");
  await expect(number).toHaveText("1");
  // One undo brings the text back under the same number
  await pressModKey(page, "z");
  await expect
    .poll(() => noteText(page, "1"))
    .toContain("A footnote near the top");
  await expect(number).toHaveText("1");
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

test("types inside an endnote at the end of the document", async ({ page }) => {
  await openHarness(page, "notes");

  // The endnote stands pages below the text that calls it, so the press on its number is also
  // what scrolls the reader to it
  const scrolled = await page.evaluate(() => window.scrollY);
  await reference(page, "1", "Endnote").click();
  await untilNoteHoldsTheCaret(page);
  expect(await page.evaluate(() => window.scrollY)).toBeGreaterThan(scrolled);

  await page.keyboard.type(" Written in place.");

  await expect
    .poll(() => noteText(page, "1", "endnote"))
    .toContain("Written in place.");
  expect(await docText(page)).not.toContain("Written in place.");
  // It is drawn after the last paragraph rather than under the sheet
  const endnotes = page.getByRole("region", {
    name: /^Endnotes on page \d+$/,
  });
  await expect(endnotes).toContainText("Written in place.");
  const area = await endnotes.boundingBox();
  const sheet = await page.locator(`.${editorClassNames.sheet}`).boundingBox();
  if (!area || !sheet) throw new Error("the endnotes were not drawn");
  expect(area.y + area.height).toBeLessThanOrEqual(sheet.y + sheet.height + 1);
});

test("takes the caret back to the reference when an endnote's number is pressed", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await reference(page, "1", "Endnote").click();
  await untilNoteHoldsTheCaret(page);

  await openNote(page).locator(`sup.${editorClassNames.noteMark}`).click();

  await expect(openNote(page)).toHaveCount(0);
  // The caret stands just past the reference, pages above where the note is drawn
  await page.keyboard.type("!");
  await expect
    .poll(() => docText(page))
    .toContain(
      "Paragraph 141 keeps the text running down the page.\n!\nParagraph 142"
    );
  await expect(reference(page, "1", "Endnote")).toBeInViewport();
});

test("inserts an endnote on the endnote key and puts the caret in it", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await page.locator(`.${editorClassNames.sheet} p`).first().click();

  // macOS keeps Command+Option+D for the Dock, so the editor binds Word for Mac's key there
  const onMac = process.platform === "darwin";
  await page.keyboard.press(
    `${onMac ? "Meta" : "Control"}+Alt+${onMac ? "e" : "d"}`
  );

  await untilNoteHoldsTheCaret(page);
  await page.keyboard.type("An endnote written where it was added.");
  await expect
    .poll(() => noteText(page, "2", "endnote"))
    .toContain("An endnote written where it was added.");
  // It is the first endnote of the document now, so the one the file arrived with moves up
  await expect(reference(page, "1", "Endnote")).toHaveCount(1);
  await expect(reference(page, "2", "Endnote")).toHaveCount(1);
  await expect(
    page.getByRole("region", { name: /^Endnotes on page \d+$/ }).first()
  ).toContainText("An endnote written where it was added.");
});
