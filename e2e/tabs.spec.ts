import { expect, type Page, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  blocks,
  caretAt,
  firstTextParagraph,
  openHarness,
  selection,
  selectText,
  settle,
} from "./support/harness";

async function followingTextLeft(page: Page, suffix: string): Promise<number> {
  return page.evaluate(
    ({ paragraphClass, wanted }) => {
      for (const paragraph of document.querySelectorAll(
        `p.${paragraphClass}`
      )) {
        const walker = document.createTreeWalker(
          paragraph,
          NodeFilter.SHOW_TEXT
        );
        let node = walker.nextNode();
        while (node) {
          const value = node.textContent ?? "";
          const at = value.indexOf(wanted);
          if (at >= 0) {
            const range = document.createRange();
            range.setStart(node, at);
            range.setEnd(node, at + 1);
            return range.getBoundingClientRect().left;
          }
          node = walker.nextNode();
        }
      }
      throw new Error("the text following the inserted tabs was not drawn");
    },
    { paragraphClass: editorClassNames.paragraph, wanted: suffix }
  );
}

interface TextBox {
  left: number;
  right: number;
  top: number;
}

async function textBox(
  page: Page,
  sourceId: number,
  wanted: string
): Promise<TextBox> {
  return page.evaluate(
    ({ paragraphClass, id, text }) => {
      const paragraph = document.querySelector(
        `p.${paragraphClass}[data-src="${id}"]`
      );
      if (!(paragraph instanceof HTMLElement)) {
        throw new Error(`paragraph ${id} was not drawn`);
      }
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
      let node = walker.nextNode();
      while (node) {
        const value = node.textContent ?? "";
        const at = value.indexOf(text);
        if (at >= 0) {
          const range = document.createRange();
          range.setStart(node, at);
          range.setEnd(node, at + text.length);
          const rect = range.getBoundingClientRect();
          return { left: rect.left, right: rect.right, top: rect.top };
        }
        node = walker.nextNode();
      }
      throw new Error(`text was not drawn in paragraph ${id}`);
    },
    {
      paragraphClass: editorClassNames.paragraph,
      id: sourceId,
      text: wanted,
    }
  );
}

async function paragraphOrigin(page: Page, sourceId: number): Promise<number> {
  return page.evaluate(
    ({ paragraphClass, id }) => {
      const paragraph = document.querySelector(
        `p.${paragraphClass}[data-src="${id}"]`
      );
      if (!(paragraph instanceof HTMLElement)) {
        throw new Error(`paragraph ${id} was not drawn`);
      }
      const style = getComputedStyle(paragraph);
      return (
        paragraph.getBoundingClientRect().left -
        Number.parseFloat(style.marginLeft || "0")
      );
    },
    { paragraphClass: editorClassNames.paragraph, id: sourceId }
  );
}

async function paragraphEndOrigin(
  page: Page,
  sourceId: number
): Promise<number> {
  return page.evaluate(
    ({ paragraphClass, id }) => {
      const paragraph = document.querySelector(
        `p.${paragraphClass}[data-src="${id}"]`
      );
      if (!(paragraph instanceof HTMLElement)) {
        throw new Error(`paragraph ${id} was not drawn`);
      }
      const style = getComputedStyle(paragraph);
      return (
        paragraph.getBoundingClientRect().right +
        Number.parseFloat(style.marginRight || "0")
      );
    },
    { paragraphClass: editorClassNames.paragraph, id: sourceId }
  );
}

function near(actual: number, expected: number): void {
  expect(Math.abs(actual - expected)).toBeLessThan(1);
}

test("successive Tab presses advance the following text independently", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const before = await blocks(page);
  const paragraph = firstTextParagraph(before);
  const suffix = paragraph.docText.slice(1);
  await caretAt(page, paragraph.index, 1);

  const positions: number[] = [];
  for (let count = 1; count <= 3; count += 1) {
    await page.keyboard.press("Tab");
    await settle(page);
    positions.push(await followingTextLeft(page, suffix));
    const after = await blocks(page);
    expect(after[paragraph.index]?.docText).toBe(
      `${paragraph.docText.slice(0, 1)}${"\t".repeat(count)}${suffix}`
    );
    await expect(page.locator(`.${editorClassNames.tabSlot}`)).toHaveCount(
      count
    );
  }

  expect(positions[1]).toBeGreaterThan((positions[0] ?? 0) + 1);
  expect(positions[2]).toBeGreaterThan((positions[1] ?? 0) + 1);
});

test("an inserted tab keeps its caret on the requested boundary", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const paragraph = firstTextParagraph(await blocks(page));
  await caretAt(page, paragraph.index, 1);
  await page.keyboard.press("Tab");
  await settle(page);

  const slot = page.locator(`.${editorClassNames.tabSlot}`).first();
  const tab = await slot.boundingBox();
  if (!tab) throw new Error("the inserted tab was not drawn");
  const caret = page.locator(`.${editorClassNames.tabCaret}`);
  near((await caret.boundingBox())?.x ?? Number.NaN, tab.x + tab.width);

  await page.keyboard.press("ArrowLeft");
  await settle(page);
  near((await caret.boundingBox())?.x ?? Number.NaN, tab.x);

  await page.keyboard.press("ArrowRight");
  await settle(page);
  near((await caret.boundingBox())?.x ?? Number.NaN, tab.x + tab.width);
});

test("a tab caret follows document zoom changes", async ({ page }) => {
  await openHarness(page, "tabs");
  await caretAt(page, 0, 1);
  await settle(page);

  await page.getByLabel("Zoom").selectOption("1.5");
  await settle(page);

  const tab = await page
    .locator(`p[data-src="0"] .${editorClassNames.tabSlot}`)
    .boundingBox();
  const caret = await page
    .locator(`.${editorClassNames.tabCaret}`)
    .boundingBox();
  if (!tab || !caret) throw new Error("the zoomed tab caret was not drawn");
  near(caret.x, tab.x);
});

test("the document's automatic tab interval controls browser tab stops", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const paragraph = page.locator(`p.${editorClassNames.paragraph}`).first();
  const origin = await paragraph.evaluate(
    (element) => element.getBoundingClientRect().left
  );
  const following = await followingTextLeft(page, "Default");

  // settings.xml declares 48pt, which Chrome draws as 64 CSS pixels.
  expect(Math.abs(following - (origin + 64))).toBeLessThan(1);
});

test("clicking and dragging a tab uses its text boundaries", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const tab = page.locator(`.${editorClassNames.tabSlot}`).first();
  await expect(tab).toHaveCSS("overflow", "clip");
  const box = await tab.boundingBox();
  if (!box) throw new Error("tab was not drawn");
  const y = box.y + box.height / 2;

  await page.mouse.click(box.x + box.width * 0.25, y);
  expect(await selection(page)).toEqual({ from: 2, to: 2, anchor: 2, head: 2 });

  await page.mouse.click(box.x + box.width * 0.75, y);
  expect(await selection(page)).toEqual({ from: 3, to: 3, anchor: 3, head: 3 });

  await page.mouse.move(box.x + box.width * 0.25, y);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width * 0.75, y);
  await page.mouse.up();
  expect(await selection(page)).toEqual({ from: 2, to: 3, anchor: 2, head: 3 });
});

test("losing the window cancels an active tab drag", async ({ page }) => {
  await openHarness(page, "tabs");
  const tab = page.locator(`.${editorClassNames.tabSlot}`).first();
  const box = await tab.boundingBox();
  if (!box) throw new Error("tab was not drawn");
  const y = box.y + box.height / 2;

  await page.mouse.move(box.x + box.width * 0.25, y);
  await page.mouse.down();
  expect(await selection(page)).toEqual({ from: 2, to: 2, anchor: 2, head: 2 });

  await page.evaluate(() => window.dispatchEvent(new Event("blur")));
  await page.mouse.move(box.x + box.width * 0.75, y);
  expect(await selection(page)).toEqual({ from: 2, to: 2, anchor: 2, head: 2 });
  await page.mouse.up();
});

test("arrow keys and Shift select a tab as one text character", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  await caretAt(page, 0, 2);

  await page.keyboard.press("ArrowLeft");
  expect(await selection(page)).toEqual({ from: 2, to: 2, anchor: 2, head: 2 });
  await page.keyboard.press("Shift+ArrowRight");
  expect(await selection(page)).toEqual({ from: 2, to: 3, anchor: 2, head: 3 });
});

test("carets remain on both sides of successive tabs", async ({ page }) => {
  await openHarness(page, "tabs");
  const paragraphIndex = 17;
  const tabs = page.locator(
    `p[data-src="${paragraphIndex}"] .${editorClassNames.tabSlot}`
  );
  await expect(tabs).toHaveCount(3);
  const tabEdges = await tabs.evaluateAll((elements) =>
    elements.map((element) => {
      const rect = element.getBoundingClientRect();
      return { left: rect.left, right: rect.right };
    })
  );
  for (const offset of [1, 2, 3, 4, 5, 6]) {
    await caretAt(page, paragraphIndex, offset);
    await settle(page);
    const caret = await page
      .locator(`.${editorClassNames.tabCaret}`)
      .boundingBox();
    if (!caret) throw new Error("the tab caret was not drawn");
    const tab = tabEdges[Math.floor((offset - 1) / 2)];
    if (!tab) throw new Error("the tab was not drawn");
    near(caret.x, offset % 2 === 1 ? tab.left : tab.right);
  }
});

test("carets use each boundary between adjacent tabs", async ({ page }) => {
  await openHarness(page, "tabs");
  const paragraphIndex = 18;
  const tabs = page.locator(
    `p[data-src="${paragraphIndex}"] .${editorClassNames.tabSlot}`
  );
  await expect(tabs).toHaveCount(4);
  const tabEdges = await tabs.evaluateAll((elements) =>
    elements.map((element) => {
      const rect = element.getBoundingClientRect();
      return { left: rect.left, right: rect.right };
    })
  );
  for (let offset = 0; offset <= 4; offset += 1) {
    await caretAt(page, paragraphIndex, offset);
    await settle(page);
    const caret = await page
      .locator(`.${editorClassNames.tabCaret}`)
      .boundingBox();
    if (!caret) throw new Error("the tab caret was not drawn");
    const expected = offset === 4 ? tabEdges[3]?.right : tabEdges[offset]?.left;
    if (expected === undefined) throw new Error("the tab was not drawn");
    near(caret.x, expected);
  }
});

test("character formatting changes a selected tab's line box", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const slot = page.locator(`p[data-src="0"] .${editorClassNames.tabSlot}`);
  const before = await slot.boundingBox();
  if (!before) throw new Error("the tab was not drawn");

  await selectText(page, 0, 1, 1);
  await page.getByLabel("Font size").selectOption("36");
  await settle(page);

  await expect(slot).toHaveCSS("font-size", "48px");
  const after = await slot.boundingBox();
  expect(after?.height).toBeGreaterThan(before.height);
});

test("a commented tab paints its annotation at text height", async ({
  page,
}) => {
  await openHarness(page, "demo");
  const paragraph = (await blocks(page)).find((block) =>
    block.docText.includes("commented range")
  );
  if (!paragraph) throw new Error("the commented range was not found");
  await caretAt(page, paragraph.index, 24);
  await page.keyboard.press("Tab");
  await settle(page);

  const boxes = await page.locator(`p[data-src="${paragraph.index}"]`).evaluate(
    (element, classes) => {
      const text = element.querySelector<HTMLElement>(
        `.${classes.comment}:not(.${classes.slot})`
      );
      const tab = element.querySelector<HTMLElement>(`.${classes.tab}`);
      const slot = element.querySelector<HTMLElement>(`.${classes.slot}`);
      if (!text || !tab || !slot) {
        throw new Error("the commented tab was not drawn");
      }
      const textRect = text.getBoundingClientRect();
      const tabRect = tab.getBoundingClientRect();
      const slotStyle = getComputedStyle(slot);
      return {
        text: {
          top: textRect.top,
          bottom: textRect.bottom,
          height: textRect.height,
        },
        tab: {
          top: tabRect.top,
          bottom: tabRect.bottom,
          height: tabRect.height,
        },
        slotBackground: slotStyle.backgroundColor,
        slotShadow: slotStyle.boxShadow,
      };
    },
    {
      comment: editorClassNames.commentRange,
      slot: editorClassNames.tabSlot,
      tab: editorClassNames.tab,
    }
  );
  near(boxes.tab.top, boxes.text.top);
  near(boxes.tab.bottom, boxes.text.bottom);
  near(boxes.tab.height, boxes.text.height);
  expect(boxes.slotBackground).toBe("rgba(0, 0, 0, 0)");
  expect(boxes.slotShadow).toBe("none");
});

test("successive selected tabs expose continuous selection geometry", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const paragraphIndex = 18;
  const paragraph = page.locator(`p[data-src="${paragraphIndex}"]`);
  const tabs = paragraph.locator(`.${editorClassNames.tabSlot}`);
  await expect(tabs).toHaveCount(4);
  const tabRects = await tabs.evaluateAll((elements) =>
    elements.map((element) => {
      const rect = element.getBoundingClientRect();
      return { left: rect.left, right: rect.right };
    })
  );
  await page.evaluate(
    (index) => window.docxHarness.selectText(index, 0, 4),
    paragraphIndex
  );
  const selected = await selection(page);
  expect(selected.to - selected.from).toBe(4);
  const selectionRects = await page.evaluate(() =>
    Array.from(getSelection()?.getRangeAt(0).getClientRects() ?? []).map(
      (rect) => ({ left: rect.left, right: rect.right })
    )
  );
  const distinctRects = selectionRects
    .filter(
      (rect, index, rects) =>
        rects.findIndex(
          (candidate) =>
            candidate.left === rect.left && candidate.right === rect.right
        ) === index
    )
    .sort((left, right) => left.left - right.left);
  expect(distinctRects.length).toBeGreaterThan(0);
  const firstSelection = distinctRects[0];
  const lastSelection = distinctRects.at(-1);
  const firstTab = tabRects[0];
  const lastTab = tabRects.at(-1);
  if (!firstSelection || !lastSelection || !firstTab || !lastTab) {
    throw new Error("the selection was not drawn");
  }
  near(firstSelection.left, firstTab.left);
  near(lastSelection.right, lastTab.right);
  for (let index = 1; index < distinctRects.length; index += 1) {
    const previous = distinctRects[index - 1];
    const current = distinctRects[index];
    if (!previous || !current) throw new Error("the selection was not drawn");
    near(previous.right, current.left);
  }
});

test("custom tab alignments use their OOXML positions", async ({ page }) => {
  await openHarness(page, "tabs");
  const origin = await paragraphOrigin(page, 1);
  const at72pt = origin + 96;
  const at144pt = origin + 192;

  near((await textBox(page, 1, "Start")).left, at72pt);
  const centered = await textBox(page, 2, "Center");
  near((centered.left + centered.right) / 2, at144pt);
  near((await textBox(page, 3, "End")).right, at144pt);
  near((await textBox(page, 4, ".")).left, at144pt);
  near((await textBox(page, 5, "List")).left, at72pt);
  near((await textBox(page, 6, "Bar skipped")).left, at72pt);
});

test("font loading completion invalidates measured tab widths", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const origin = await paragraphOrigin(page, 2);

  await page.locator('p[data-src="2"]').evaluate((paragraph) => {
    if (!(paragraph instanceof HTMLElement)) {
      throw new Error("the centered paragraph was not drawn");
    }
    paragraph.style.fontFamily = "monospace";
    document.fonts.dispatchEvent(new Event("loadingdone"));
  });
  await settle(page);

  const centered = await textBox(page, 2, "Center");
  near((centered.left + centered.right) / 2, origin + 192);
});

test("each line resolves the paragraph's tab stops independently", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const origin = await paragraphOrigin(page, 7);
  const first = await textBox(page, 7, "One");
  const second = await textBox(page, 7, "Two");
  near(first.left, origin + 96);
  near(second.left, origin + 96);
  expect(second.top).toBeGreaterThan(first.top);
});

test("a tab inside a link keeps its line height, caret baseline, and underline", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const paragraph = page.locator(
    `p[data-src="8"].${editorClassNames.paragraph}`
  );
  const before = await paragraph.boundingBox();
  await caretAt(page, 8, 8);
  await page.keyboard.press("Tab");
  await settle(page);

  const report = await blocks(page);
  expect(report[8]?.docText).toBe("https://\texample.com");
  const slot = page.locator(
    `p[data-src="8"] .${editorClassNames.link} .${editorClassNames.tabSlot}`
  );
  await expect(slot).toHaveCount(1);
  await expect(
    page.locator(`p[data-src="8"] .${editorClassNames.link}`)
  ).toHaveCSS("border-bottom-style", "solid");
  const slotBox = await slot.boundingBox();
  expect(slotBox?.width).toBeGreaterThan(1);
  const after = await paragraph.boundingBox();
  if (!before || !after || !slotBox) {
    throw new Error("the linked tab was not drawn");
  }
  near(after.height, before.height);

  await caretAt(page, 8, 8);
  const tabCaret = await page.evaluate(() => window.docxHarness.caretBox());
  await caretAt(page, 8, 9);
  const textCaret = await page.evaluate(() => window.docxHarness.caretBox());
  near(tabCaret.top, textCaret.top);
  near(tabCaret.bottom, textCaret.bottom);
});

test("start and end stops follow a right-to-left paragraph's logical edges", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const startOrigin = await paragraphEndOrigin(page, 9);
  const endOrigin = await paragraphEndOrigin(page, 10);

  near((await textBox(page, 9, "אבג")).right, startOrigin - 96);
  near((await textBox(page, 10, "אבג")).left, endOrigin - 192);
});

test("automatic tabs use the page-margin grid after paragraph indentation", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const origin = await paragraphOrigin(page, 11);

  near((await textBox(page, 11, "Indented")).left, origin + 64);
});

test("an automatic tab moved by soft wrapping is recalculated on its new line", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const origin = await paragraphOrigin(page, 12);
  const firstLine = await textBox(page, 12, "W");
  const following = await textBox(page, 12, "Wrapped");

  expect(following.top).toBeGreaterThan(firstLine.top);
  near(following.left, origin + 64);
});

test("decimal tabs use the following run's locale separator", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const englishOrigin = await paragraphOrigin(page, 13);
  const germanOrigin = await paragraphOrigin(page, 14);

  near((await textBox(page, 13, ".")).left, englishOrigin + 192);
  near((await textBox(page, 14, ",")).left, germanOrigin + 192);
});

test("Strict start indentation follows the leading edge of an RTL paragraph", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const origin = await paragraphEndOrigin(page, 15);

  near((await textBox(page, 15, "א")).right, origin - 40);
  near((await textBox(page, 15, "אבג")).right, origin - 64);
});

test("Transitional left indentation follows the leading edge of an RTL paragraph", async ({
  page,
}) => {
  await openHarness(page, "tabs");
  const origin = await paragraphEndOrigin(page, 16);

  near((await textBox(page, 16, "א")).right, origin - 40);
  near((await textBox(page, 16, "אבג")).right, origin - 64);
});
