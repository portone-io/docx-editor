/** Real table boxes and collapsed borders are required to verify row-boundary pagination. */

import { expect, test } from "@playwright/test";
import { editorAttributes, editorClassNames } from "../src/styles/classNames";
import { openHarness, settle, tableRows } from "./support/harness";

test("a long table continues between rows and repeats its header", async ({
  page,
}) => {
  await openHarness(page, "long-table");
  await settle(page);

  const table = page.locator(`.${editorClassNames.table}`);
  const originalRows = table.locator(
    `:scope > tbody > .${editorClassNames.tableRow}:not([${editorAttributes.tableRepeatedHeader}])`
  );
  const spacers = table.locator(
    `:scope > tbody > [${editorAttributes.tablePageSpace}]`
  );
  const repeatedHeaders = table.locator(
    `:scope > tbody > [${editorAttributes.tableRepeatedHeader}]`
  );

  expect(await tableRows(page)).toBe(41);
  await expect(originalRows).toHaveCount(41);
  expect(await spacers.count()).toBeGreaterThanOrEqual(2);
  await expect(repeatedHeaders).toHaveCount(await spacers.count());
  await expect(repeatedHeaders.first()).toHaveText("HeadingValue");
  await expect(repeatedHeaders.first()).toHaveAttribute("aria-hidden", "true");
  await expect(repeatedHeaders.first()).toHaveAttribute(
    "contenteditable",
    "false"
  );
  const settledSpacers = await spacers.evaluateAll((elements) =>
    elements.map((element) => element.getAttribute("data-table-page-space"))
  );
  await settle(page);
  await settle(page);
  expect(
    await spacers.evaluateAll((elements) =>
      elements.map((element) => element.getAttribute("data-table-page-space"))
    )
  ).toEqual(settledSpacers);

  const pageSplits = page.locator(`.${editorClassNames.pageSplit}`);
  expect(await pageSplits.count()).toBe(await spacers.count());
  const splitTops = await pageSplits.evaluateAll((elements) =>
    elements.map((element) => element.getBoundingClientRect().top)
  );
  const spacerBoxes = await spacers.evaluateAll((elements) =>
    elements.map((element) => {
      const box = element.getBoundingClientRect();
      return { top: box.top, bottom: box.bottom };
    })
  );
  const rowBoxes = await originalRows.evaluateAll((elements) =>
    elements.map((element) => {
      const box = element.getBoundingClientRect();
      return { top: box.top, bottom: box.bottom };
    })
  );

  splitTops.forEach((top, index) => {
    const spacer = spacerBoxes[index];
    expect(spacer).toBeDefined();
    expect(top).toBeGreaterThanOrEqual((spacer?.top ?? 0) - 1);
    expect(top).toBeLessThanOrEqual((spacer?.bottom ?? 0) + 1);
    expect(
      rowBoxes.some((row) => row.top < top - 1 && row.bottom > top + 1)
    ).toBe(false);
  });
});
