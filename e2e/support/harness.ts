/**
 * Opening the harness page and reading it back.
 *
 * Everything a test asserts is read through `window.docxHarness` (`../harness/api`), so a test
 * never has to know how the editor draws itself.
 */

import { expect, type Page } from "@playwright/test";
import { editorClassNames } from "../../src/styles/classNames";
import type {
  BlockReport,
  CaretBox,
  CompositionCounts,
  SelectionReport,
} from "../harness/api";

export type { BlockReport, CaretBox, CompositionCounts };

/** Opens the page over one fixture and waits until the editor has drawn it */
export async function openHarness(page: Page, fixture: string): Promise<void> {
  await page.goto(`/?fixture=${fixture}`);
  await page.waitForFunction(() => Boolean(window.docxHarness));
  await expect(page.locator(`.${editorClassNames.sheet}`)).toBeVisible();
  // The page overlay finds its positions on the frames after the first paint
  await settle(page);
}

/** Waits for two animation frames, which is where a remeasure lands */
export function settle(page: Page): Promise<void> {
  return page.evaluate(
    () =>
      new Promise<void>((resolve) => {
        requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
      })
  );
}

export function blocks(page: Page): Promise<BlockReport[]> {
  return page.evaluate(() => window.docxHarness.blocks());
}

export function docText(page: Page): Promise<string> {
  return page.evaluate(() => window.docxHarness.text());
}

export function composing(page: Page): Promise<boolean> {
  return page.evaluate(() => window.docxHarness.composing());
}

export function compositions(page: Page): Promise<CompositionCounts> {
  return page.evaluate(() => window.docxHarness.compositions());
}

export function pushes(page: Page): Promise<string> {
  return page.evaluate(() => window.docxHarness.pushes());
}

export function pagination(page: Page): Promise<string> {
  return page.evaluate(() => window.docxHarness.pagination());
}

export function spaces(page: Page): Promise<string> {
  return page.evaluate(() => window.docxHarness.spaces());
}

export function caretBox(page: Page): Promise<CaretBox> {
  return page.evaluate(() => window.docxHarness.caretBox());
}

export function selection(page: Page): Promise<SelectionReport> {
  return page.evaluate(() => window.docxHarness.selection());
}

export function lockedText(page: Page): Promise<string> {
  return page.evaluate(() => window.docxHarness.lockedText());
}

export function blockHeight(page: Page, blockIndex: number): Promise<number> {
  return page.evaluate(
    (index) => window.docxHarness.blockHeight(index),
    blockIndex
  );
}

export function caretAt(
  page: Page,
  blockIndex: number,
  offset: number
): Promise<number> {
  return page.evaluate(
    ([block, at]) => window.docxHarness.caretAt(block, at),
    [blockIndex, offset]
  );
}

export function selectText(
  page: Page,
  blockIndex: number,
  offset: number,
  length: number
): Promise<void> {
  return page.evaluate(
    ([block, at, span]) => window.docxHarness.selectText(block, at, span),
    [blockIndex, offset, length]
  );
}

export function caretInCell(page: Page): Promise<string> {
  return page.evaluate(() => window.docxHarness.caretInCell());
}

export function rightClick(page: Page): Promise<void> {
  return page.evaluate(() => {
    window.docxHarness.rightClick();
  });
}

export function tableRows(page: Page): Promise<number> {
  return page.evaluate(() => window.docxHarness.tableRows());
}

/** What the focus is on, by its accessible name or the text written on it */
export function focused(page: Page): Promise<string> {
  return page.evaluate(() => {
    const found = document.activeElement;
    if (!(found instanceof HTMLElement)) return "";
    return found.getAttribute("aria-label") ?? found.textContent ?? "";
  });
}

export function lock(
  page: Page,
  blockIndex: number,
  offset: number,
  length: number
): Promise<boolean> {
  return page.evaluate(
    ([block, at, span]) => window.docxHarness.lock(block, at, span),
    [blockIndex, offset, length]
  );
}

/** The block of a fixture the tests type into: the first paragraph carrying text */
export function firstTextParagraph(found: readonly BlockReport[]): BlockReport {
  const paragraph = found.find(
    (block) => block.type === "paragraph" && block.docText.length > 0
  );
  if (!paragraph) throw new Error("the fixture holds no paragraph with text");
  return paragraph;
}
