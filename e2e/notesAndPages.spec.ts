import { expect, type Locator, type Page, test } from "@playwright/test";
import { editorAttributes, editorClassNames } from "../src/styles/classNames";
import { openHarness, pressModKey } from "./support/harness";

async function boxOf(locator: Locator) {
  const box = await locator.boundingBox();
  if (!box) throw new Error("the element was not drawn");
  return box;
}

/** Where each gap between two pages starts on the screen, top to bottom */
function pageEnds(page: Page): Promise<number[]> {
  return page
    .locator(`.${editorClassNames.pageSplit}`)
    .evaluateAll((bands) =>
      bands.map((band) => band.getBoundingClientRect().top)
    );
}

/** The page a point on the screen stands on, counted from one */
function pageAt(ends: readonly number[], y: number): number {
  return ends.filter((end) => end <= y).length + 1;
}

test("the demo starts at page one, draws its footnote on the paper, and keeps its closing order", async ({
  page,
}) => {
  await openHarness(page, "demo");

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
  await expect(
    page.getByRole("region", { name: /^Footnotes on page \d+$/ })
  ).toContainText("A footnote appears at the bottom of its page in Word.");
  await expect(
    page.getByRole("region", { name: /^Endnotes on page \d+$/ })
  ).toContainText(
    "An endnote is collected at the end of a document or section in Word."
  );

  const section13 = await page
    .getByText("13. Cell alignment and padding", { exact: true })
    .boundingBox();
  const imageLocator = page.locator(`.${editorClassNames.imageBox}`);
  const image = await imageLocator.boundingBox();
  if (!section13 || !image)
    throw new Error("the closing demo blocks were not drawn");
  expect(image.y).toBeGreaterThan(section13.y);

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

  const section14 = page.getByText("14. Comments, bookmarks and notes", {
    exact: true,
  });
  await expect(
    section14.locator(
      `xpath=ancestor::p//*[@${editorAttributes.breakType}="page"]`
    )
  ).toHaveCount(0);
});

test("draws a footnote at the foot of the page that refers to it, above the footer", async ({
  page,
}) => {
  await openHarness(page, "notes");
  const areas = page.getByRole("region", { name: /^Footnotes on page \d+$/ });
  await expect(areas).toHaveCount(2);

  const reference = page.locator(
    `.${editorClassNames.sheet} [aria-label="Footnote 2"]`
  );
  const referenceBox = await boxOf(reference);
  const ends = await pageEnds(page);
  const referencePage = pageAt(ends, referenceBox.y);
  expect(referencePage).toBeGreaterThan(1);

  const area = page.getByRole("region", {
    name: `Footnotes on page ${referencePage}`,
  });
  const areaBox = await boxOf(area);
  expect(areaBox.y).toBeGreaterThan(referenceBox.y + referenceBox.height);
  expect(pageAt(ends, areaBox.y + areaBox.height - 1)).toBe(referencePage);

  // Every line of text on that page ends above the footnotes kept at its foot
  const lastLineBottom = await page
    .locator(`.${editorClassNames.sheet} > p`)
    .evaluateAll(
      (paragraphs, { top, bottom }) =>
        Math.max(
          ...paragraphs
            .map((paragraph) => paragraph.getBoundingClientRect())
            .filter((rect) => rect.top >= top && rect.top < bottom)
            .map((rect) => rect.bottom)
        ),
      {
        top: ends[referencePage - 2] ?? 0,
        bottom: areaBox.y + areaBox.height,
      }
    );
  expect(lastLineBottom).toBeLessThanOrEqual(areaBox.y + 1);

  const footer = page
    .locator(`.${editorClassNames.pageFooter}`)
    .nth(referencePage - 1);
  expect(areaBox.y + areaBox.height).toBeLessThanOrEqual(
    (await boxOf(footer)).y + 1
  );

  await expect(area.locator(`sup.${editorClassNames.noteMark}`)).toHaveText(
    "2"
  );
  await expect(area.getByText("bold words")).toHaveCSS("font-weight", "700");
  await expect(area).toContainText("and a second paragraph.");
  await expect(area).not.toContainText("near the top of the document");
  // The room kept is the height the footnote was measured at, so none of it is hidden or cut off
  await expect(area.locator(`.${editorClassNames.noteRow}`)).toBeVisible();
  expect(
    await area.evaluate(
      (element) => element.scrollHeight - element.clientHeight
    )
  ).toBeLessThanOrEqual(1);
  const rule = await boxOf(area.locator(`.${editorClassNames.noteSeparator}`));
  expect(rule.width).toBeLessThan(areaBox.width / 2);
});

test("copies a selection inside a footnote area with none of the document's own markup", async ({
  page,
  context,
}) => {
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  await openHarness(page, "notes");
  const area = page
    .getByRole("region", { name: /^Footnotes on page \d+$/ })
    .last();
  await expect(area).toContainText("bold words");

  // The area is not the editor, so the browser writes the copy itself from what is selected
  await area.evaluate((element) => {
    const range = document.createRange();
    range.selectNodeContents(element);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  });
  await pressModKey(page, "c");
  const html = await page.evaluate(async () => {
    const items = await navigator.clipboard.read();
    const item = items.find((entry) => entry.types.includes("text/html"));
    return item === undefined ? "" : (await item.getType("text/html")).text();
  });

  expect(html).toContain("bold words");
  expect(html).not.toContain("<w:");
  expect(html).not.toMatch(/data-(rpr|rattrs|fmt|ppr|pattrs|src|xml)=/);
});

test("draws the endnotes after the last paragraph, on the paper", async ({
  page,
}) => {
  await openHarness(page, "notes");

  const sheetBox = await boxOf(page.locator(`.${editorClassNames.sheet}`));
  const area = page.getByRole("region", { name: /^Endnotes on page \d+$/ });
  const areaBox = await boxOf(area);
  // On the paper rather than under it, and below every line of the body
  expect(areaBox.y + areaBox.height).toBeLessThanOrEqual(
    sheetBox.y + sheetBox.height + 1
  );
  const lastLineBottom = await page
    .locator(`.${editorClassNames.sheet} > p`)
    .evaluateAll((paragraphs) =>
      Math.max(
        ...paragraphs.map(
          (paragraph) => paragraph.getBoundingClientRect().bottom
        )
      )
    );
  expect(areaBox.y).toBeGreaterThanOrEqual(lastLineBottom - 1);

  await expect(area.locator(`sup.${editorClassNames.noteMark}`)).toHaveText(
    "1"
  );
  await expect(area.getByText("An endnote in italics.")).toHaveCSS(
    "font-style",
    "italic"
  );
  await expect(area).not.toContainText("bold words");
  const rule = await boxOf(area.locator(`.${editorClassNames.noteSeparator}`));
  expect(rule.width).toBeLessThan(areaBox.width / 2);
});

test("a link the document underlines itself is drawn with one line", async ({
  page,
}) => {
  await openHarness(page, "demo");

  const link = page.locator(`.${editorClassNames.link}`).first();
  await expect(link.locator(`.${editorClassNames.run}`)).toHaveCSS(
    "text-decoration-line",
    "underline"
  );
  await expect(link).toHaveCSS("border-bottom-style", "none");
});
