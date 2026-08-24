import { expect, type Page, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  blocks,
  openHarness,
  rightClick,
  selectText,
  settle,
} from "./support/harness";

async function layoutSnapshot(page: Page) {
  return page.evaluate((classes) => {
    const layer = document.querySelector(`.${classes.pageLayer}`);
    if (!(layer instanceof HTMLElement)) throw new Error("page layer missing");
    return {
      sheetHeight: layer.style.getPropertyValue("--docx-editor-sheet-height"),
      pages: document.querySelectorAll(`.${classes.pageBadge}`).length,
    };
  }, editorClassNames);
}

async function frameSnapshots(page: Page, count: number) {
  return page.evaluate(
    async ({ classes, frames }) => {
      const layer = document.querySelector(`.${classes.pageLayer}`);
      if (!(layer instanceof HTMLElement))
        throw new Error("page layer missing");
      const found: string[] = [];
      for (let frame = 0; frame < frames; frame += 1) {
        await new Promise<void>((resolve) =>
          requestAnimationFrame(() => resolve())
        );
        found.push(
          `${layer.style.getPropertyValue("--docx-editor-sheet-height")}/${document.querySelectorAll(`.${classes.pageBadge}`).length}`
        );
      }
      return found;
    },
    { classes: editorClassNames, frames: count }
  );
}

test("the demo keeps its pagination stable in narrow layouts", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1000, height: 900 });
  await openHarness(page, "demo");
  await page.getByRole("button", { name: "Show comments" }).click();
  await settle(page);
  const baseline = await layoutSnapshot(page);
  const zoom = page.getByLabel("Zoom");
  await expect(zoom).toHaveValue("fit-width");
  const font = page.getByLabel("Font", { exact: true });
  await expect(font).toHaveCSS("width", "112px");
  await expect(font).toHaveCSS("text-overflow", "ellipsis");
  await expect(font).toHaveCSS("white-space", "nowrap");
  for (const size of [
    {
      width: 839,
      commentWidth: 240,
      canvasPadding: "10px 12px",
      cardPadding: 10,
      metaFontSize: 11,
      bodyFontSize: 13,
    },
    {
      width: 719,
      commentWidth: 200,
      canvasPadding: "8px",
      cardPadding: 8,
      metaFontSize: 10,
      bodyFontSize: 12,
    },
    {
      width: 559,
      commentWidth: 180,
      canvasPadding: "6px",
      cardPadding: 6,
      metaFontSize: 10,
      bodyFontSize: 12,
    },
    {
      width: 419,
      commentWidth: 180,
      canvasPadding: "6px",
      cardPadding: 6,
      metaFontSize: 10,
      bodyFontSize: 12,
    },
  ]) {
    await page.setViewportSize({ width: size.width, height: 900 });
    await settle(page);
    await settle(page);

    expect(await layoutSnapshot(page)).toEqual(baseline);
    expect(new Set(await frameSnapshots(page, 12)).size).toBe(1);
    await expect(page.locator(`.${editorClassNames.commentsPanel}`)).toHaveCSS(
      "width",
      `${size.commentWidth}px`
    );
    await expect(page.locator(`.${editorClassNames.commentsCanvas}`)).toHaveCSS(
      "padding",
      size.canvasPadding
    );
    await expect(
      page.locator(`.${editorClassNames.commentCard}`).first()
    ).toHaveCSS("padding", `${size.cardPadding}px`);
    const editor = page.locator(`.${editorClassNames.root}`);
    await editor.evaluate((element) => {
      element.scrollLeft = element.scrollWidth;
    });
    await settle(page);
    const editorBox = await editor.boundingBox();
    expect(
      await editor.evaluate((element) => {
        if (!(element instanceof HTMLElement)) {
          throw new Error("responsive editor is not an HTML element");
        }
        return element.offsetWidth - element.clientWidth;
      })
    ).toBeGreaterThanOrEqual(10);
    const commentsBox = await page
      .locator(`.${editorClassNames.commentsPanel}`)
      .boundingBox();
    if (!editorBox || !commentsBox) {
      throw new Error("responsive editor layout missing");
    }
    expect(editorBox.x + editorBox.width).toBeLessThanOrEqual(
      commentsBox.x + 1
    );
    await expect(
      page.locator(`.${editorClassNames.commentMeta}`).first()
    ).toHaveCSS("font-size", `${size.metaFontSize}px`);
    await expect(
      page.locator(`.${editorClassNames.commentBody}`).first()
    ).toHaveCSS("font-size", `${size.bodyFontSize}px`);
  }

  await zoom.selectOption("1");
  await settle(page);
  await settle(page);
  await expect(zoom).toHaveValue("1");
  await expect(page.locator(`.${editorClassNames.pageLayer}`)).toHaveCSS(
    "zoom",
    "1"
  );
  expect(await layoutSnapshot(page)).toEqual(baseline);
  expect(new Set(await frameSnapshots(page, 12)).size).toBe(1);
  expect(
    await page
      .locator(`.${editorClassNames.root}`)
      .evaluate((root) => root.scrollWidth > root.clientWidth)
  ).toBe(true);

  await zoom.selectOption("fit-width");
  await settle(page);
  await settle(page);
  await expect(page.locator(`.${editorClassNames.pageLayer}`)).toHaveCSS(
    "zoom",
    "0.5"
  );

  const commentHeader = page
    .locator(`.${editorClassNames.commentHeader}`)
    .first();
  await expect(commentHeader).toHaveCSS("flex-direction", "column");
  const author = await commentHeader
    .locator(`.${editorClassNames.commentAuthor}`)
    .boundingBox();
  const date = await commentHeader
    .locator(`.${editorClassNames.commentDate}`)
    .boundingBox();
  const icons = await commentHeader
    .locator(`.${editorClassNames.commentIconActions}`)
    .boundingBox();
  const firstIcon = await commentHeader
    .locator(`.${editorClassNames.commentIconActions} button`)
    .first()
    .boundingBox();
  if (!author || !date || !icons || !firstIcon) {
    throw new Error("responsive comment header missing");
  }
  expect(date.y).toBeGreaterThan(author.y);
  expect(icons.y).toBeGreaterThan(date.y);
  expect(firstIcon.width).toBe(24);
  expect(firstIcon.height).toBe(24);

  await page.getByRole("button", { name: "Text color" }).click();
  const panel = page.locator(`.${editorClassNames.popover}`);
  await expect(panel).toHaveCSS("font-size", "11px");
  const box = await panel.boundingBox();
  if (!box) throw new Error("popover missing");
  expect(box.x).toBeGreaterThanOrEqual(8);
  expect(box.x + box.width).toBeLessThanOrEqual(411);
  expect(box.y).toBeGreaterThanOrEqual(8);
  expect(box.y + box.height).toBeLessThanOrEqual(892);
  expect(
    await panel.evaluate((element) => ({
      horizontal: element.scrollWidth > element.clientWidth,
      vertical: element.scrollHeight > element.clientHeight,
    }))
  ).toEqual({ horizontal: false, vertical: false });
  await page.keyboard.press("Escape");

  const target = (await blocks(page)).find(
    (block) => block.type === "paragraph" && block.docText.length > 6
  );
  if (!target) throw new Error("demo paragraph missing");
  await selectText(page, target.index, 1, 5);
  await rightClick(page);
  const menu = page.getByRole("menu", { name: "Text actions" });
  await expect(menu).toHaveCSS("font-size", "11px");
  const menuBox = await menu.boundingBox();
  const menuIcon = await menu.locator("svg").first().boundingBox();
  if (!menuBox || !menuIcon) throw new Error("text menu missing");
  expect(menuBox.x).toBeGreaterThanOrEqual(8);
  expect(menuBox.x + menuBox.width).toBeLessThanOrEqual(411);
  expect(menuBox.y).toBeGreaterThanOrEqual(8);
  expect(menuBox.y + menuBox.height).toBeLessThanOrEqual(892);
  expect(menuIcon.width).toBe(13);
  expect(menuIcon.height).toBe(13);
});

test("opening and resizing a narrow comment rail preserves left scroll", async ({
  page,
}) => {
  await page.setViewportSize({ width: 419, height: 900 });
  await openHarness(page, "demo");
  const editor = page.locator(`.${editorClassNames.root}`);
  await expect(
    page.locator(`.${editorClassNames.commentsPanel}`)
  ).toHaveAttribute("data-view", "rail");
  await expect
    .poll(() => editor.evaluate((element) => element.scrollLeft))
    .toBe(0);

  await page.setViewportSize({ width: 1000, height: 900 });
  await page.getByLabel("Zoom").selectOption("1");
  await settle(page);
  await settle(page);
  await editor.evaluate((element) => {
    element.scrollLeft = 0;
  });

  await page.setViewportSize({ width: 419, height: 900 });
  await settle(page);
  await settle(page);

  expect(
    await editor.evaluate(
      (element) => element.scrollWidth > element.clientWidth
    )
  ).toBe(true);
  await expect
    .poll(() => editor.evaluate((element) => element.scrollLeft))
    .toBe(0);
});
