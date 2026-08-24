import { expect, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import { blocks, openHarness, selectText, settle } from "./support/harness";

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
  const reopenedCardBox = await card.boundingBox();
  if (!reopenedCardBox) throw new Error("the reopened comment was not drawn");
  expect(Math.abs(reopenedCardBox.y - rangeBox.y)).toBeLessThan(40);
});
