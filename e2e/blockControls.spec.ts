/**
 * The edge of a block-level content control under a real browser.
 *
 * A refused keystroke has two halves: the editor builds no change, and the browser makes none of
 * its own in the contenteditable. Only the first is visible to jsdom, so the second is here: the
 * key is pressed for real and the document is read back out of the model afterwards.
 */

import { expect, test } from "@playwright/test";
import {
  type BlockReport,
  blocks,
  caretAt,
  docText,
  openHarness,
  selection,
} from "./support/harness";

const FIXTURE = "content-controls";

/** The first block-level control of the fixture, which holds two paragraphs of its own */
function firstControl(found: readonly BlockReport[]): BlockReport {
  const control = found.find((block) => block.type === "sdtBlock");
  if (!control) throw new Error("the fixture holds no block control");
  return control;
}

/**
 * The caret offset of the end of a control holding two paragraphs.
 * The text reads them joined by a line, where the document puts two positions - one closing the
 * first paragraph, one opening the second - between them.
 */
function endOfTwoParagraphs(control: BlockReport): number {
  const lines = control.docText.split("\n");
  if (lines.length !== 2)
    throw new Error("the control no longer holds two paragraphs");
  return control.docText.length + 1;
}

test.describe("editing at a block control's edge", () => {
  test("Backspace at the start of its first block changes nothing", async ({
    page,
  }) => {
    await openHarness(page, FIXTURE);
    const control = firstControl(await blocks(page));
    const before = await docText(page);
    const at = await caretAt(page, control.index, 0);

    await page.keyboard.press("Backspace");

    expect(await docText(page)).toBe(before);
    expect(await selection(page)).toEqual({
      from: at,
      to: at,
      anchor: at,
      head: at,
    });
  });

  test("Delete at the end of its last block changes nothing", async ({
    page,
  }) => {
    await openHarness(page, FIXTURE);
    const control = firstControl(await blocks(page));
    const before = await docText(page);
    const at = await caretAt(page, control.index, endOfTwoParagraphs(control));

    await page.keyboard.press("Delete");

    expect(await docText(page)).toBe(before);
    expect(await selection(page)).toEqual({
      from: at,
      to: at,
      anchor: at,
      head: at,
    });
  });
});
