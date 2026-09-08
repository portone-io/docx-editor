import { expect, test } from "@playwright/test";
import { openHarness } from "./support/harness";

test("numbering-level formatting belongs to the marker, and suffixes control its spacing", async ({
  page,
}) => {
  await openHarness(page, "list-definitions");
  const roman = page.locator('p[data-marker="I."]');
  const drawn = await roman.evaluate((p) => {
    const marker = getComputedStyle(p, "::before");
    const text = p.querySelector("span") ?? p;
    const body = getComputedStyle(text);
    return {
      markerWeight: marker.fontWeight,
      markerColor: marker.color,
      markerSize: parseFloat(marker.fontSize),
      bodyWeight: body.fontWeight,
      bodyColor: body.color,
      width: parseFloat(marker.minWidth),
    };
  });
  expect(drawn.markerWeight).toBe("700");
  expect(drawn.markerColor).toBe("rgb(31, 56, 100)");
  expect(drawn.markerSize).toBeCloseTo((14 * 96) / 72, 2);
  expect(drawn.bodyWeight).not.toBe("700");
  expect(drawn.bodyColor).not.toBe(drawn.markerColor);
  expect(drawn.width).toBe(24);
  for (const marker of ["一 ", "일)"]) {
    const spacing = await page
      .locator(`p[data-marker="${marker}"]`)
      .evaluate((p) => {
        const style = getComputedStyle(p, "::before");
        return [style.minWidth, style.paddingRight];
      });
    expect(spacing).toEqual(["0px", "0px"]);
  }
});
