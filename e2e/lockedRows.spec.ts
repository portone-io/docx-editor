/**
 * A table row a template locked, in a real browser: what typing and deleting there leaves behind,
 * what the application is told about it, and how the row is drawn.
 *
 * The row's cells state their shading as `auto`, which the editor draws as a transparent inline
 * background. The lock is drawn as a layer over that background rather than as a background colour
 * of its own, which a colour stated inline would win over; the computed style is what shows it.
 */

import { expect, type Page, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  caretInText,
  docText,
  openHarness,
  refusals,
  rightClick,
  settle,
} from "./support/harness";

/** What the browser paints behind the cell or control holding this text, and its outline */
function paintOf(
  page: Page,
  needle: string,
  selector: string
): Promise<{ image: string; outline: string; cursor: string }> {
  return page.evaluate(
    ([text, css]) => {
      const holder = Array.from(document.querySelectorAll(css)).find(
        (element) => element.textContent?.includes(text)
      );
      if (!holder) throw new Error(`nothing drawn holds ${text}`);
      const style = getComputedStyle(holder);
      return {
        image: style.backgroundImage,
        outline: `${style.outlineStyle} ${style.outlineColor}`,
        cursor: style.cursor,
      };
    },
    [needle, selector] as const
  );
}

function setFrameProperty(page: Page, name: string, value: string) {
  return page.evaluate(
    ([frame, property, set]) => {
      const found = document.querySelector<HTMLElement>(`.${frame}`);
      found?.style.setProperty(property, set);
    },
    [editorClassNames.frame, name, value] as const
  );
}

test.beforeEach(async ({ page }) => {
  await openHarness(page, "locked-rows");
});

test("typing and deleting in a locked row changes nothing and says which row refused", async ({
  page,
}) => {
  const before = await docText(page);

  await caretInText(page, "Per session", 3);
  await page.keyboard.type("x");
  await page.keyboard.press("Backspace");
  await page.keyboard.press("Delete");
  await settle(page);

  expect(await docText(page)).toBe(before);
  const told = await refusals(page);
  expect(told.map((refusal) => refusal.action)).toEqual([
    "insert",
    "delete",
    "delete",
  ]);
  for (const refusal of told) {
    expect(refusal.reason).toBe("lock");
    expect(refusal.controls[0]).toMatchObject({
      tag: "PRICE_ROWS",
      alias: "PRICE_ROWS",
      id: 21,
      lock: "sdtContentLocked",
      level: "row",
    });
  }
});

test("typing in a row nobody locked goes through and reports nothing", async ({
  page,
}) => {
  await caretInText(page, "Agreed", 0);
  await page.keyboard.type("Y");
  await settle(page);

  expect(await docText(page)).toContain("YAgreed");
  expect(await refusals(page)).toEqual([]);
});

test("the right-click menu over a locked row says why its rows stand disabled", async ({
  page,
}) => {
  await caretInText(page, "Per item", 2);
  await rightClick(page);

  const menu = page.getByRole("menu", { name: "Table actions" });
  await expect(menu).toBeVisible();
  await expect(menu).toHaveAccessibleDescription(
    "Locked content can't be edited."
  );
  await expect(
    menu.getByRole("menuitem", { name: "Delete row" })
  ).toHaveAttribute("aria-disabled", "true");
});

test("a locked row is drawn over the cells' own transparent shading", async ({
  page,
}) => {
  const cell = `td.${editorClassNames.tableCell}`;
  const locked = await paintOf(page, "Per session", cell);
  expect(locked.image).toContain("rgba(255, 235, 59, 0.5)");
  expect(locked.outline).toBe("solid rgb(249, 168, 37)");
  expect(locked.cursor).toBe("not-allowed");

  // A row nobody locked is not drawn at all
  const open = await paintOf(page, "Agreed", cell);
  expect(open.image).toBe("none");
  expect(open.outline).toContain("none");
});

test("the host themes shut and open content through custom properties", async ({
  page,
}) => {
  await setFrameProperty(
    page,
    "--docx-editor-locked-background",
    "rgb(240 240 240)"
  );
  await setFrameProperty(
    page,
    "--docx-editor-locked-outline",
    "2px solid rgb(10 20 30)"
  );
  await setFrameProperty(
    page,
    "--docx-editor-control-background",
    "rgb(255 245 157)"
  );

  const cell = `td.${editorClassNames.tableCell}`;
  const row = await paintOf(page, "Per session", cell);
  expect(row.image).toContain("rgb(240, 240, 240)");
  expect(row.outline).toBe("solid rgb(10, 20, 30)");

  const field = await paintOf(page, "Sample Name", `.${editorClassNames.sdt}`);
  expect(field.image).toContain("rgb(240, 240, 240)");

  const openField = await paintOf(
    page,
    "Open note",
    `.${editorClassNames.sdt}`
  );
  expect(openField.image).toContain("rgb(255, 245, 157)");
});
