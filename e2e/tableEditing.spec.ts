import { expect, test } from "@playwright/test";
import { editorAttributes } from "../src/styles/classNames";
import {
  blocks,
  caretAt,
  firstTextParagraph,
  openHarness,
  pagination,
  settle,
} from "./support/harness";

test("inserted tables keep an editable paragraph between them", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink");
  const target = firstTextParagraph(await blocks(page));
  await caretAt(page, target.index, 0);

  const insertTable = page.getByRole("button", { name: "Insert table" });
  await insertTable.click();
  await page.getByRole("gridcell", { name: "1 by 1 table" }).click();

  const separatorIndex = target.index + 2;
  await caretAt(page, separatorIndex, 0);
  await page.keyboard.press("Backspace");
  expect((await blocks(page))[separatorIndex]).toMatchObject({
    type: "paragraph",
    docText: "",
  });

  await insertTable.click();
  await page.getByRole("gridcell", { name: "1 by 1 table" }).click();
  expect(
    (await blocks(page))
      .slice(target.index + 1, target.index + 5)
      .map(({ type }) => type)
  ).toEqual(["table", "paragraph", "table", "paragraph"]);
});

test("row resizing previews height before committing OOXML", async ({
  page,
}) => {
  await openHarness(page, "long-table");
  const row = page.getByRole("row", { name: "Row 1 Value 1", exact: true });
  await row.scrollIntoViewIfNeeded();
  const cell = row.getByRole("cell").first();
  const before = await row.boundingBox();
  const edge = await cell.boundingBox();
  if (!before || !edge) throw new Error("table row was not drawn");
  const beforePagination = await pagination(page);
  const beforeTrPr = await row.getAttribute("data-trpr");

  await page.mouse.move(edge.x + edge.width / 2, edge.y + edge.height - 1);
  await page.mouse.down();
  await page.mouse.move(edge.x + edge.width / 2, edge.y + edge.height + 23);
  const preview = await row.boundingBox();
  expect(preview?.height ?? 0).toBeGreaterThan(before.height + 15);
  await expect(page.locator(".docx-editor-row-guide")).toHaveCount(1);
  await settle(page);
  expect(await pagination(page)).toBe(beforePagination);
  expect(await row.getAttribute("data-trpr")).toBe(beforeTrPr);
  await page.mouse.up();

  await expect(row).toHaveAttribute(
    "data-trpr",
    /w:trHeight w:val="(?!780")[0-9]+" w:hRule="exact"/
  );
  const after = await row.boundingBox();
  expect(after?.height ?? 0).toBeGreaterThan(before.height + 15);
  await expect.poll(() => pagination(page)).not.toBe(beforePagination);
});

test("a row preview follows the pointer across a fixed page boundary", async ({
  page,
}) => {
  await openHarness(page, "long-table");
  const spacer = page.locator(`tr[${editorAttributes.tablePageSpace}]`).first();
  const boundaryRow = spacer.locator(
    `xpath=preceding-sibling::tr[not(@${editorAttributes.tablePageSpace}) and ` +
      `not(@${editorAttributes.tableRepeatedHeader})][1]`
  );
  const rowLabel = await boundaryRow.getByRole("cell").first().innerText();
  const row = page
    .getByText(rowLabel, { exact: true })
    .locator("xpath=ancestor::tr");
  await row.scrollIntoViewIfNeeded();
  const cell = row.getByRole("cell").first();
  const edge = await cell.boundingBox();
  if (!edge) throw new Error("table row was not drawn");
  const beforePagination = await pagination(page);
  const pointerY = edge.y + edge.height + 120;

  await page.mouse.move(edge.x + edge.width / 2, edge.y + edge.height - 1);
  await page.mouse.down();
  await page.mouse.move(edge.x + edge.width / 2, pointerY);
  await settle(page);
  const heldPagination = await pagination(page);
  const preview = await row.boundingBox();
  if (!preview) throw new Error("table row preview was not drawn");

  const samples = await row.evaluate(async (element, tablePageSpace) => {
    const frames: string[] = [];
    for (let index = 0; index < 12; index += 1) {
      await new Promise<void>((resolve) =>
        requestAnimationFrame(() => resolve())
      );
      const box = element.getBoundingClientRect();
      const spaces = Array.from(
        document.querySelectorAll<HTMLElement>(`tr[${tablePageSpace}]`)
      ).map((space) => space.getAttribute(tablePageSpace));
      frames.push(JSON.stringify([box.top, box.height, spaces]));
    }
    return frames;
  }, editorAttributes.tablePageSpace);

  expect(new Set(samples).size).toBe(1);
  expect(Math.abs(preview.y + preview.height - pointerY)).toBeLessThan(3);
  expect(heldPagination).toBe(beforePagination);
  expect(await pagination(page)).toBe(beforePagination);
  await page.mouse.up();
});
