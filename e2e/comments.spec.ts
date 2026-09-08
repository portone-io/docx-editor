import { expect, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  blocks,
  openHarness,
  rightClick,
  selectText,
  settle,
} from "./support/harness";

test("the all-comments panel scrolls itself before continuing through the document", async ({
  page,
}) => {
  await openHarness(page, "demo");
  await page.getByTestId("editor").evaluate((element) => {
    element.style.height = "320px";
  });
  await page.getByRole("button", { name: "Show comments" }).click();
  await settle(page);

  const panel = page.locator(`.${editorClassNames.commentsPanel}`);
  const editor = page.locator(`.${editorClassNames.root}`);
  await editor.evaluate((element) => {
    element.scrollLeft = element.scrollWidth;
  });
  await settle(page);
  await panel.hover();
  await page.mouse.wheel(0, 120);
  await expect
    .poll(() => editor.evaluate((element) => element.scrollTop))
    .toBeGreaterThan(0);

  await editor.evaluate((element) => {
    element.scrollTop = 0;
  });
  await panel
    .locator(`.${editorClassNames.commentsCanvas}`)
    .evaluate((element) => {
      element.style.minHeight = "1200px";
    });
  await panel.hover();
  await page.mouse.wheel(0, 120);
  await expect
    .poll(() => panel.evaluate((element) => element.scrollTop))
    .toBeGreaterThan(0);

  await panel.evaluate((element) => {
    element.scrollTop = element.scrollHeight;
  });
  const documentBefore = await editor.evaluate((element) => element.scrollTop);
  await page.mouse.wheel(0, 120);
  await expect
    .poll(() => editor.evaluate((element) => element.scrollTop))
    .toBeGreaterThan(documentBefore);
});

test("a comment starts from the selected text menu and stands beside its anchor", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const targets = (await blocks(page)).filter(
    (block) => block.type === "paragraph" && block.docText.length > 20
  );
  const target = targets[0];
  if (!target) throw new Error("the fixture holds no paragraph long enough");

  await selectText(page, target.index, 2, 6);
  await page.evaluate(() => window.docxHarness.rightClick());
  await page.getByRole("menuitem", { name: "Add comment" }).click();
  await page.getByRole("textbox", { name: "Comment text" }).fill("Review this");
  await page.getByRole("button", { name: "Comment", exact: true }).click();
  await settle(page);

  const range = page.locator(`.${editorClassNames.commentRange}`).first();
  const card = page.locator(`.${editorClassNames.commentCard}`).first();
  const sheet = page.locator(`.${editorClassNames.sheet}`);
  await expect(card).toContainText("Review this");
  const rangeBox = await range.boundingBox();
  const cardBox = await card.boundingBox();
  const sheetBox = await sheet.boundingBox();
  if (!rangeBox || !cardBox || !sheetBox) {
    throw new Error("the comment was not drawn");
  }
  expect(Math.abs(cardBox.y - rangeBox.y)).toBeLessThan(40);
  expect(cardBox.x - (sheetBox.x + sheetBox.width)).toBeGreaterThan(0);
  expect(cardBox.x - (sheetBox.x + sheetBox.width)).toBeLessThan(48);

  await page.getByTestId("editor").evaluate((element) => {
    element.style.height = "500px";
  });
  await settle(page);
  const editor = page.locator(`.${editorClassNames.root}`);
  const scrollBefore = await editor.evaluate((element) => element.scrollTop);
  await card.hover();
  await page.mouse.wheel(0, 180);
  await expect
    .poll(() => editor.evaluate((element) => element.scrollTop))
    .toBeGreaterThan(scrollBefore);

  const horizontalBeforeResize = await editor.evaluate(
    (element) => element.scrollLeft
  );
  await page.setViewportSize({ width: 900, height: 720 });
  await settle(page);
  const resizedSheetBox = await sheet.boundingBox();
  const resizedCardBox = await card.boundingBox();
  if (!resizedSheetBox || !resizedCardBox) {
    throw new Error("the resized comment rail was not drawn");
  }
  expect(
    resizedCardBox.x - (resizedSheetBox.x + resizedSheetBox.width)
  ).toBeGreaterThan(0);
  expect(
    resizedCardBox.x - (resizedSheetBox.x + resizedSheetBox.width)
  ).toBeLessThan(48);
  expect(await editor.evaluate((element) => element.scrollLeft)).toBe(
    horizontalBeforeResize
  );
  await editor.evaluate((element) => {
    element.scrollLeft = element.scrollWidth;
  });
  await settle(page);
  const narrowSheetBox = await sheet.boundingBox();
  const narrowCardBox = await card.boundingBox();
  if (!narrowSheetBox || !narrowCardBox) {
    throw new Error("the narrow comment rail was not drawn");
  }
  expect(
    narrowCardBox.x - (narrowSheetBox.x + narrowSheetBox.width)
  ).toBeGreaterThan(0);
  expect(
    narrowCardBox.x - (narrowSheetBox.x + narrowSheetBox.width)
  ).toBeLessThan(48);
  const narrowEditorBox = await editor.boundingBox();
  if (!narrowEditorBox) throw new Error("the narrow editor was not drawn");
  expect(narrowCardBox.x + narrowCardBox.width).toBeLessThanOrEqual(
    narrowEditorBox.x + narrowEditorBox.width - 15
  );

  const secondTarget = targets[1];
  if (!secondTarget) throw new Error("the fixture holds no second paragraph");
  await selectText(page, secondTarget.index, 2, 6);
  await page.evaluate(() => window.docxHarness.rightClick());
  await page.getByRole("menuitem", { name: "Add comment" }).click();
  await page
    .getByRole("textbox", { name: "Comment text" })
    .fill("Second review");
  await page.getByRole("button", { name: "Comment", exact: true }).click();
  await page
    .getByRole("button", { name: "Reply to comment: Review this" })
    .click();
  await settle(page);
  const positioned = page.locator(`.${editorClassNames.commentPosition}`);
  const firstPosition = await positioned.nth(0).boundingBox();
  const secondPosition = await positioned.nth(1).boundingBox();
  if (!firstPosition || !secondPosition) {
    throw new Error("the comment cards were not drawn");
  }
  expect(secondPosition.y).toBeGreaterThanOrEqual(
    firstPosition.y + firstPosition.height
  );
  await page.getByRole("button", { name: "Cancel" }).click();
  await page.getByRole("button", { name: "Delete" }).last().click();

  await page.getByRole("button", { name: "Resolve" }).click();
  await expect(
    page.getByRole("complementary", { name: "Comments" })
  ).toHaveCount(0);
  await page.getByRole("button", { name: "Show comments" }).click();
  await expect(
    page.getByRole("complementary", { name: "Comments" })
  ).toHaveAttribute("data-view", "all");
  await page.getByRole("button", { name: "Reopen" }).click();
  await expect(card).toContainText("Review this");
  await page.getByRole("button", { name: "Hide comments" }).click();
  await expect(
    page.getByRole("complementary", { name: "Comments" })
  ).toHaveAttribute("data-view", "rail");
  await expect(card).toContainText("Review this");
  // Measured again rather than against the rectangle taken at the top of this test: the page has
  // been scrolled several times since, and where the card stands is a question about the anchor
  // as it stands now. The rail places itself over the two frames after the panel is put away, so
  // the distance is polled rather than read once
  await expect
    .poll(async () => {
      const cardNow = await card.boundingBox();
      const rangeNow = await range.boundingBox();
      if (!cardNow || !rangeNow) return Number.POSITIVE_INFINITY;
      return Math.abs(cardNow.y - rangeNow.y);
    })
    .toBeLessThan(40);
});

/**
 * Opening the composer used to take the reader back to the top of the document: the form is put
 * where its anchor is only once the comment rail has measured itself, and focusing it before then
 * had the browser scroll to where it stood in the meantime, which was the first page.
 */
test("the composer leaves the page where the reader left it", async ({
  page,
}) => {
  await openHarness(page, "demo");
  await page.getByTestId("editor").evaluate((element) => {
    element.style.height = "500px";
  });
  await settle(page);

  const editor = page.locator(`.${editorClassNames.root}`);
  const paragraphs = (await blocks(page)).filter(
    (block) => block.type === "paragraph" && block.docText.length > 20
  );
  const target = paragraphs[paragraphs.length - 1];
  if (!target) throw new Error("the fixture holds no paragraph long enough");

  await selectText(page, target.index, 2, 6);
  await editor.evaluate((element) => {
    element.scrollTop = element.scrollHeight;
  });
  await settle(page);
  const before = await editor.evaluate((element) => element.scrollTop);
  expect(before).toBeGreaterThan(100);

  await rightClick(page);
  await page.getByRole("menuitem", { name: "Add comment" }).click();
  await settle(page);
  await expect(
    page.getByRole("textbox", { name: "Comment text" })
  ).toBeFocused();

  expect(await editor.evaluate((element) => element.scrollTop)).toBe(before);
});
