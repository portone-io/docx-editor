import { expect, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  blocks,
  docText,
  openHarness,
  selectText,
  settle,
} from "./support/harness";

/**
 * A commenter's surface in a real browser: the text stays what it was under the keys, the drag
 * selection the browser makes still reaches the editor, and the comment goes onto that selection
 * through the same menu an editor uses.
 */
test("a commenter selects text and comments on it, and typing changes nothing", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink", "comment");
  const before = await docText(page);
  const targets = (await blocks(page)).filter(
    (block) => block.type === "paragraph" && block.docText.length > 20
  );
  const target = targets[0];
  if (!target) throw new Error("the fixture holds no paragraph long enough");

  await expect(page.getByRole("toolbar")).toHaveCount(0);

  await selectText(page, target.index, 2, 6);
  await page.keyboard.type("typed");
  await page.keyboard.press("Backspace");
  expect(await docText(page)).toBe(before);

  await page.evaluate(() => window.docxHarness.rightClick());
  const rows = page.getByRole("menuitem");
  await expect(rows).toHaveText([/^Copy/, "Add comment"]);
  await rows.filter({ hasText: "Add comment" }).click();
  await page.getByRole("textbox", { name: "Comment text" }).fill("Review this");
  await page.getByRole("button", { name: "Comment", exact: true }).click();
  await settle(page);

  await expect(
    page.locator(`.${editorClassNames.commentCard}`).first()
  ).toContainText("Review this");
  // The harness text breaks at the comment's markers, so the paragraph is read off the page
  await expect(
    page.locator("p").filter({ hasText: target.docText.slice(0, 30) })
  ).toHaveText(target.docText);
});

/**
 * Every entry a commenter is offered is about the selected text, so a right click landing outside
 * one is worth more to the reader as the browser's own menu than as a menu of dead rows.
 */
test("a commenter right clicking with nothing selected keeps the browser's own menu", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink", "comment");
  const target = (await blocks(page)).find(
    (block) => block.type === "paragraph" && block.docText.length > 20
  );
  if (!target) throw new Error("the fixture holds no paragraph long enough");

  await selectText(page, target.index, 2, 0);
  await page.evaluate(() => window.docxHarness.rightClick());
  await expect(page.getByRole("menuitem")).toHaveCount(0);
});

test("a reader is handed the browser's own menu and no comment composer", async ({
  page,
}) => {
  await openHarness(page, "kitchen-sink", "readOnly");
  const targets = (await blocks(page)).filter(
    (block) => block.type === "paragraph" && block.docText.length > 20
  );
  const target = targets[0];
  if (!target) throw new Error("the fixture holds no paragraph long enough");

  await selectText(page, target.index, 2, 6);
  await page.evaluate(() => window.docxHarness.rightClick());
  await expect(page.getByRole("menuitem")).toHaveCount(0);
});
