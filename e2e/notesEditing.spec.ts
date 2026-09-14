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
  noteOpening,
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

async function enterNote(page: Page, label: string, kind = "Footnote") {
  await reference(page, label, kind).click();
  await untilNoteHoldsTheCaret(page);
}

type NoteKind = "footnote" | "endnote";

/** One note of the fixture as the gestures below reach it, and as a test reads its story back */
interface NoteUnderTest {
  readonly kind: NoteKind;
  /** The id the notes part names the entry by */
  readonly id: string;
  /** The number the body draws its reference as, which is how a test presses it */
  readonly label: string;
  /** What the reference is called on the page, which names the kind */
  readonly called: string;
}

const FOOTNOTE: NoteUnderTest = {
  kind: "footnote",
  id: "1",
  label: "1",
  called: "Footnote",
};
const ENDNOTE: NoteUnderTest = {
  kind: "endnote",
  id: "1",
  label: "1",
  called: "Endnote",
};
const NOTES: readonly NoteUnderTest[] = [FOOTNOTE, ENDNOTE];

/**
 * Whether the browser's caret stands at the first place in the note a caret may: just after the
 * number, which is where the note's own text begins.
 *
 * It is read off the browser's own selection rather than off the editor, because what the gestures
 * below are about is where the browser will write.
 */
function caretIsAtTheHead(page: Page): Promise<boolean> {
  return page.evaluate((classes) => {
    const line = document.querySelector(
      `.${classes.noteRowOpen} .ProseMirror p`
    );
    const number = line?.querySelector(`sup.${classes.noteMark}`);
    const selection = window.getSelection();
    const at = selection?.focusNode ?? null;
    if (!line || !number || at === null || selection?.isCollapsed !== true) {
      return false;
    }
    const before = document.createRange();
    before.setStart(line, 0);
    before.setEnd(at, selection.focusOffset);
    return before.toString() === (number.textContent ?? "");
  }, editorClassNames);
}

/**
 * Walks the caret back to the head of the note, which is where a reader who means to write in
 * front of the number ends up.
 *
 * Home is not the key for it: on a Mac it scrolls the page instead. The walk is driven one press
 * at a time and read back after each, since Chrome drops an arrow press that arrives while the
 * main thread is busy (`./README.md`), and it stops the moment the caret arrives rather than
 * counting characters, so the step that lands on the number is never taken by accident.
 */
async function caretToTheHead(page: Page, spelled: string): Promise<void> {
  for (let step = 0; step < spelled.length + 10; step += 1) {
    if (await caretIsAtTheHead(page)) return;
    await page.keyboard.press("ArrowLeft");
    await page.waitForTimeout(8);
  }
  throw new Error("the caret never reached the head of the note");
}

/**
 * The same walk and one step more, which is the reader asking for the place in front of the
 * number. Nothing they write from there may land there, and that is what each gesture below
 * measures: a walk that stopped at the head would pass whether the rule held or not.
 */
async function caretBeforeTheNumber(
  page: Page,
  spelled: string
): Promise<void> {
  await caretToTheHead(page, spelled);
  await page.keyboard.press("ArrowLeft");
  await page.waitForTimeout(8);
}

/**
 * Whether the browser's own caret stands in front of the number, which is the place a press may
 * not leave it: the editor can say the caret is elsewhere and the browser still write here.
 */
function caretInFrontOfTheNumber(page: Page): Promise<boolean> {
  return page.evaluate((classes) => {
    const line = document.querySelector(
      `.${classes.noteRowOpen} .ProseMirror p`
    );
    const number = line?.querySelector(`sup.${classes.noteMark}`);
    const selection = window.getSelection();
    const at = selection?.focusNode ?? null;
    if (!line || !number || at === null || selection?.isCollapsed !== true) {
      return false;
    }
    const before = document.createRange();
    before.setStart(line, 0);
    before.setEnd(at, selection.focusOffset);
    return before.toString() === "";
  }, editorClassNames);
}

/** What the open note draws in front of the number it is drawn by */
function drawnBeforeTheNumber(page: Page): Promise<string> {
  return page.evaluate((classes) => {
    const line = document.querySelector(
      `.${classes.noteRowOpen} .ProseMirror p`
    );
    const number = line?.querySelector(`sup.${classes.noteMark}`);
    if (!line || !number) throw new Error("the open note draws no number");
    const range = document.createRange();
    range.setStart(line, 0);
    range.setEndBefore(number);
    return range.toString();
  }, editorClassNames);
}

/**
 * Nothing stands in front of the number, in the story the file is written from or on the screen.
 *
 * The story is read for the node the number is, since the number spells no text of its own and a
 * test reading the note as text could not tell where it stands.
 */
async function numberStaysFirst(
  page: Page,
  note: NoteUnderTest
): Promise<void> {
  expect((await noteOpening(page, note.id, note.kind))[0]).toBe(
    "rawRunContent"
  );
  expect(await drawnBeforeTheNumber(page)).toBe("");
}

/** Hands the view a clipboard the way the browser does for a paste of plain text */
function pasteText(page: Page, text: string): Promise<void> {
  return page.evaluate((pasted) => {
    const data = new DataTransfer();
    data.setData("text/plain", pasted);
    document.activeElement?.dispatchEvent(
      new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData: data,
      })
    );
  }, text);
}

test("types, formats, and undoes inside a footnote at the foot of its page", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await enterNote(page, "1");

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
  await enterNote(page, "2");
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
 * Every gesture that puts a reader at the head of a note, in both kinds of note.
 *
 * The number Word draws a note by is the first thing the entry holds, and the file has to go out
 * that way whatever the reader did. Two rules hold it: the caret is kept off the place in front of
 * the number, and an edit that wrote there has the number put back at the head
 * (`editor/notes/noteSurface`). Only a browser can drive the gestures that reach for that place -
 * a composition, a selection the browser itself replaces, a caret the browser left where
 * ProseMirror would not have put it - which is why they are all driven here rather than in jsdom.
 *
 * A footnote and an endnote share the surface, so each gesture is driven in both.
 */
for (const note of NOTES) {
  test(`types in front of the number of a ${note.kind} and writes after it`, async ({
    page,
  }) => {
    await openHarness(page, "notes");
    await enterNote(page, note.label, note.called);
    const spelled = await noteText(page, note.id, note.kind);

    await caretBeforeTheNumber(page, spelled);
    await page.keyboard.type("Head.");
    await settle(page);

    await expect
      .poll(() => noteText(page, note.id, note.kind))
      .toBe(`Head.${spelled}`);
    await numberStaysFirst(page, note);
  });

  test(`composes in front of the number of a ${note.kind} and writes after it`, async ({
    page,
  }) => {
    await openHarness(page, "notes");
    await enterNote(page, note.label, note.called);
    const spelled = await noteText(page, note.id, note.kind);
    const cdp = await imeSession(page);

    await caretBeforeTheNumber(page, spelled);
    for (const stage of ["ㅇ", "아", "안"]) {
      await setComposition(cdp, stage);
      // The syllables may not be seen standing in front of the number even mid-composition
      expect(await drawnBeforeTheNumber(page), `at buffer ${stage}`).toBe("");
    }
    await commitComposition(cdp, "안");
    await settle(page);

    await expect
      .poll(() => noteText(page, note.id, note.kind))
      .toBe(`안${spelled}`);
    await numberStaysFirst(page, note);
  });

  /**
   * The reader's own report: four jamo struck at the head of a note, which a 2-beolsik IME
   * delivers as four compositions rather than one, since ㄹ closes no syllable with ㄹ after it.
   */
  test(`writes jamo in front of the number of a ${note.kind} one composition at a time`, async ({
    page,
  }) => {
    await openHarness(page, "notes");
    await enterNote(page, note.label, note.called);
    const spelled = await noteText(page, note.id, note.kind);
    const cdp = await imeSession(page);

    await caretBeforeTheNumber(page, spelled);
    for (let struck = 0; struck < 4; struck += 1) {
      await setComposition(cdp, "ㄹ");
      await commitComposition(cdp, "ㄹ");
      expect(await drawnBeforeTheNumber(page), `at jamo ${struck}`).toBe("");
    }
    await settle(page);

    await expect
      .poll(() => noteText(page, note.id, note.kind))
      .toBe(`ㄹㄹㄹㄹ${spelled}`);
    await numberStaysFirst(page, note);
  });

  test(`composes over the number a ${note.kind} is drawn by and keeps it`, async ({
    page,
  }) => {
    await openHarness(page, "notes");
    await enterNote(page, note.label, note.called);
    const spelled = await noteText(page, note.id, note.kind);
    const cdp = await imeSession(page);

    await caretToTheHead(page, spelled);
    await page.keyboard.press("Shift+ArrowLeft");
    await compose(cdp, ["ㅇ", "아", "안"]);
    await settle(page);
    await commitComposition(cdp, "안");
    await settle(page);

    await expect
      .poll(() => noteText(page, note.id, note.kind))
      .toBe(`안${spelled}`);
    await numberStaysFirst(page, note);
  });

  /**
   * Writing a note again from nothing: select all of it, then type.
   *
   * A composition replaces what is selected because the browser does the replacing, and Chrome
   * will not touch a selection that begins at an element it may not edit. A note always begins
   * with the number it is drawn by, so the whole of a note is exactly such a selection, and the
   * syllables used to land in front of the words they were meant to replace. The selection is
   * taken away before the composition opens now (`editor/plugins/compositionSelection`).
   */
  test(`writes a ${note.kind} again from nothing with hangul`, async ({
    page,
  }) => {
    await openHarness(page, "notes");
    await enterNote(page, note.label, note.called);
    const cdp = await imeSession(page);

    await pressModKey(page, "a");
    await compose(cdp, ["ㅇ", "아", "안"]);
    await settle(page);
    await commitComposition(cdp, "안");
    await settle(page);

    // Nothing of what stood there is left, and the number the note is drawn by still is
    await expect.poll(() => noteText(page, note.id, note.kind)).toBe("안");
    await numberStaysFirst(page, note);
    expect(await composing(page)).toBe(false);
  });

  /**
   * A paste that arrives mid-composition, which ProseMirror hands to the browser rather than
   * reading itself, so what lands is whatever the browser writes into the open composition.
   */
  test(`pastes while composing in front of the number of a ${note.kind}`, async ({
    page,
  }) => {
    await openHarness(page, "notes");
    await enterNote(page, note.label, note.called);
    const spelled = await noteText(page, note.id, note.kind);
    const cdp = await imeSession(page);

    await caretBeforeTheNumber(page, spelled);
    await setComposition(cdp, "ㅇ");
    await pasteText(page, "Pasted");
    await settle(page);
    await commitComposition(cdp, "안");
    await settle(page);

    await numberStaysFirst(page, note);
  });

  /**
   * The sequence a reader reported: Home, which lands on the number itself, and one press more,
   * which is the press that used to walk past it and leave the browser's caret in front of the
   * number while the editor still said otherwise. Every press that reaches that way is answered
   * where there is nothing left of the number to reach (`editor/notes/noteSurface`), so the caret
   * stays where it stands and what is typed next lands after the number.
   */
  test(`presses past the head of a ${note.kind} leave the caret after its number`, async ({
    page,
  }) => {
    await openHarness(page, "notes");
    await enterNote(page, note.label, note.called);
    const spelled = await noteText(page, note.id, note.kind);

    await caretToTheHead(page, spelled);
    for (const press of ["Home", "ArrowLeft", "Home", "ArrowUp"]) {
      await page.keyboard.press(press);
      await page.waitForTimeout(8);
      expect(await caretInFrontOfTheNumber(page), `after ${press}`).toBe(false);
    }
    await page.keyboard.type("X");
    await settle(page);

    await expect
      .poll(() => noteText(page, note.id, note.kind))
      .toBe(`X${spelled}`);
    await numberStaysFirst(page, note);
  });

  /**
   * The same with Shift held, which builds a selection rather than moving a caret. Selecting the
   * number is an ordinary thing to do, so it stays possible; what may not happen is a press
   * reaching past it, and what is typed over it puts the number back.
   */
  test(`shift presses past the head of a ${note.kind} reach no further than its number`, async ({
    page,
  }) => {
    await openHarness(page, "notes");
    await enterNote(page, note.label, note.called);
    const spelled = await noteText(page, note.id, note.kind);

    await caretToTheHead(page, spelled);
    for (const press of ["Shift+ArrowLeft", "Shift+ArrowLeft", "Shift+Home"]) {
      await page.keyboard.press(press);
      await page.waitForTimeout(8);
      expect(await caretInFrontOfTheNumber(page), `after ${press}`).toBe(false);
    }
    await page.keyboard.type("X");
    await settle(page);

    await expect
      .poll(() => noteText(page, note.id, note.kind))
      .toBe(`X${spelled}`);
    await numberStaysFirst(page, note);
  });

  /**
   * The other end of the same note, where a press has nothing to reach either. Nothing stands
   * there for a press to walk past, and this is what says so.
   */
  test(`presses past the end of a ${note.kind} write at its end`, async ({
    page,
  }) => {
    await openHarness(page, "notes");
    await enterNote(page, note.label, note.called);
    const spelled = await noteText(page, note.id, note.kind);

    for (const press of [
      "End",
      "ArrowRight",
      "ArrowRight",
      "Shift+ArrowRight",
      "Shift+End",
      "End",
    ]) {
      await page.keyboard.press(press);
      await page.waitForTimeout(8);
    }
    await page.keyboard.type("X");
    await settle(page);

    await expect
      .poll(() => noteText(page, note.id, note.kind))
      .toBe(`${spelled}X`);
    await numberStaysFirst(page, note);
  });
}

/**
 * The number Word draws a note by is a preserved chip inside the note, and a browser deletes an
 * inline atom itself: the deletion arrives as a DOM change read back rather than as a key the
 * keymap answered, which is the path only a real browser takes.
 */
test("keeps the note's own number when its whole text is deleted", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await enterNote(page, "1");
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
  await enterNote(page, "1");

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

test("inserts an endnote from the right click menu and puts the caret in it", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await page.locator(`.${editorClassNames.sheet} p`).first().click();

  await rightClick(page);
  await page.getByRole("menuitem", { name: "Insert endnote" }).click();

  await untilNoteHoldsTheCaret(page);
  await page.keyboard.type("An endnote written from the menu.");
  await expect
    .poll(() => noteText(page, "2", "endnote"))
    .toContain("An endnote written from the menu.");
  // It is the first endnote of the document now, and both stand after the last paragraph
  await expect(
    page.getByRole("region", { name: /^Endnotes on page \d+$/ }).first()
  ).toContainText("An endnote written from the menu.");
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

test("goes back to the reference when the number of a drawn note is pressed", async ({
  page,
}) => {
  await openHarness(page, "notes");
  const area = page.getByRole("region", { name: "Footnotes on page 1" });
  await expect(area).toContainText("A footnote near the top of the document.");

  await area.locator(`sup.${editorClassNames.noteMark}`).click();

  // No view is opened over it: the number is the way back, wherever the note is drawn
  await expect(openNote(page)).toHaveCount(0);
  await page.keyboard.type("!");
  await expect
    .poll(() => docText(page))
    .toContain(
      "Paragraph 5 keeps the text running down the page.\n!\nParagraph 6"
    );
});

test("goes back to the reference when the number of the open note is pressed", async ({
  page,
}) => {
  await openHarness(page, "notes");
  await enterNote(page, "1");

  await openNote(page).locator(`sup.${editorClassNames.noteMark}`).click();

  await expect(openNote(page)).toHaveCount(0);
  await page.keyboard.type("!");
  await expect
    .poll(() => docText(page))
    .toContain(
      "Paragraph 5 keeps the text running down the page.\n!\nParagraph 6"
    );
});
