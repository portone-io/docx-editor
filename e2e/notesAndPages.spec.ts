import { expect, test } from "@playwright/test";
import { editorAttributes, editorClassNames } from "../src/styles/classNames";
import { openHarness } from "./support/harness";

test("the demo aligns its notes with the paper and starts at page one", async ({
  page,
}) => {
  await openHarness(page, "demo");

  const sheet = page.locator(`.${editorClassNames.sheet}`);
  const notes = page.locator(`.${editorClassNames.notesPanel}`);
  const sheetBox = await sheet.boundingBox();
  const notesBox = await notes.boundingBox();
  if (!sheetBox || !notesBox)
    throw new Error("the paper or notes were not drawn");

  expect(Math.abs(notesBox.x - sheetBox.x)).toBeLessThan(1);
  expect(Math.abs(notesBox.width - sheetBox.width)).toBeLessThan(1);
  await expect(page.getByLabel("Style").locator("option")).toHaveText([
    "Normal",
    "Title",
    "Subtitle",
    "Heading 1",
    "Heading 2",
    "Heading 3",
  ]);
  await expect(
    page.locator(`.${editorClassNames.pageHeader}`).first()
  ).toHaveCSS("text-align", "right");
  await expect(
    page.locator(`.${editorClassNames.pageFooter}`).first()
  ).toHaveText("Page 1");

  const section14 = await page
    .getByText("14. Cell alignment and padding", { exact: true })
    .boundingBox();
  const imageLocator = page.locator(`.${editorClassNames.imageBox}`);
  const image = await imageLocator.boundingBox();
  if (!section14 || !image)
    throw new Error("the closing demo blocks were not drawn");
  expect(image.y).toBeGreaterThan(section14.y);

  const imageSpacing = await imageLocator.evaluate((element) => {
    const paragraph = element.closest("p");
    const before = paragraph?.previousElementSibling;
    const after = paragraph?.nextElementSibling;
    return {
      before: [
        before?.previousElementSibling?.textContent,
        before?.textContent,
      ],
      after: [after?.textContent, after?.nextElementSibling?.textContent],
    };
  });
  expect(imageSpacing).toEqual({ before: ["", ""], after: ["", ""] });

  const section13 = page.getByText("13. Comments, bookmarks and notes", {
    exact: true,
  });
  await expect(
    section13.locator(
      `xpath=ancestor::p//*[@${editorAttributes.breakType}="page"]`
    )
  ).toHaveCount(0);
});
