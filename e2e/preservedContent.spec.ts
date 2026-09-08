import { expect, test } from "@playwright/test";
import { openHarness, selection } from "./support/harness";

test("preserved chips can be selected and deleted while field pieces remain guarded", async ({
  page,
}) => {
  await openHarness(page, "preserved-markup");
  for (const element of ["sym", "ins"]) {
    const chip = page.locator(`[data-element="${element}"]`);
    await chip.click();
    const selected = await selection(page);
    expect(selected.to - selected.from).toBe(1);
    await page.keyboard.press("Backspace");
    await expect(chip).toHaveCount(0);
  }
  const field = page.locator('[data-element="fldChar"]').first();
  await field.click();
  const selected = await selection(page);
  expect(selected.to - selected.from).toBe(1);
  await page.keyboard.press("Backspace");
  await expect(page.locator('[data-element="fldChar"]')).toHaveCount(3);
});

test("dragging across a revision chip includes it in the text selection", async ({
  page,
}) => {
  await openHarness(page, "preserved-markup");
  const chip = page.locator('[data-element="ins"]');
  await chip.scrollIntoViewIfNeeded();
  const bounds = await chip.boundingBox();
  if (bounds === null) throw new Error("the revision chip is not drawn");
  const y = bounds.y + bounds.height / 2;
  await page.mouse.move(bounds.x - 3, y);
  await page.mouse.down();
  await page.mouse.move(bounds.x + bounds.width + 3, y, { steps: 12 });
  await page.mouse.up();
  expect(
    await page.evaluate(() => window.getSelection()?.toString())
  ).toContain("insertion");
  await page.keyboard.press("Backspace");
  await expect(chip).toHaveCount(0);
});
