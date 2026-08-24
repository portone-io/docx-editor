import { expect, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  blocks,
  firstTextParagraph,
  openHarness,
  selectText,
  settle,
} from "./support/harness";

const TINY_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNg" +
  "AAIAAAUAAen63NgAAAAASUVORK5CYII=";

test("a browser paste keeps supported font formatting", async ({ page }) => {
  await openHarness(page, "demo");
  const target = firstTextParagraph(await blocks(page));
  await selectText(page, target.index, 0, Math.min(4, target.docText.length));

  await page.evaluate((sheetClass) => {
    const data = new DataTransfer();
    data.setData("text/plain", "Styled paste");
    data.setData(
      "text/html",
      '<span style="font-family: Arial; font-size: 16pt; font-weight: 700; font-style: italic; text-decoration: underline line-through; color: #123456; background-color: #abcdef">Styled paste</span>'
    );
    const sheet = document.querySelector(`.${sheetClass}`);
    if (!(sheet instanceof HTMLElement))
      throw new Error("editor sheet missing");
    sheet.dispatchEvent(
      new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData: data,
      })
    );
  }, editorClassNames.sheet);
  await settle(page);

  const pasted = page
    .locator(`.${editorClassNames.run}`)
    .filter({ hasText: "Styled paste" })
    .first();
  await expect(pasted).toHaveCSS("font-family", /Arial/);
  await expect(pasted).toHaveCSS("font-size", "21.3333px");
  await expect(pasted).toHaveCSS("font-weight", "700");
  await expect(pasted).toHaveCSS("font-style", "italic");
  await expect(pasted).toHaveCSS(
    "text-decoration-line",
    /underline.*line-through|line-through.*underline/
  );
  await expect(pasted).toHaveCSS("color", "rgb(18, 52, 86)");
  await expect(pasted).toHaveCSS("background-color", "rgb(171, 205, 239)");
});

test("a browser paste maps a heading to the document style", async ({
  page,
}) => {
  await openHarness(page, "demo");
  const target = firstTextParagraph(await blocks(page));
  await selectText(page, target.index, 0, Math.min(4, target.docText.length));

  await page.evaluate((sheetClass) => {
    const data = new DataTransfer();
    data.setData("text/plain", "Pasted heading");
    data.setData("text/html", "<h1>Pasted heading</h1>");
    const sheet = document.querySelector(`.${sheetClass}`);
    if (!(sheet instanceof HTMLElement)) {
      throw new Error("editor sheet missing");
    }
    sheet.dispatchEvent(
      new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData: data,
      })
    );
  }, editorClassNames.sheet);
  await settle(page);

  const heading = page
    .locator(`p.${editorClassNames.paragraph}`)
    .filter({ hasText: "Pasted heading" });
  await expect(heading).toHaveAttribute("data-ppr", /w:val="Heading1"/);
  await expect(page.getByLabel("Style")).toHaveValue("id:Heading1");
});

test("a browser paste embeds an image copied from a web app", async ({
  page,
}) => {
  await page.route("https://assets.example/pasted.png", (route) =>
    route.fulfill({
      body: Buffer.from(TINY_PNG_BASE64, "base64"),
      contentType: "image/png",
      headers: { "access-control-allow-origin": "*" },
    })
  );
  await openHarness(page, "demo");
  const target = firstTextParagraph(await blocks(page));
  await selectText(page, target.index, 0, Math.min(4, target.docText.length));

  await page.evaluate((sheetClass) => {
    const data = new DataTransfer();
    data.setData("text/plain", "seal");
    data.setData(
      "text/html",
      '<img src="https://assets.example/pasted.png" alt="seal">'
    );
    const sheet = document.querySelector(`.${sheetClass}`);
    if (!(sheet instanceof HTMLElement))
      throw new Error("editor sheet missing");
    sheet.dispatchEvent(
      new ClipboardEvent("paste", {
        bubbles: true,
        cancelable: true,
        clipboardData: data,
      })
    );
  }, editorClassNames.sheet);

  const image = page.locator(`img.${editorClassNames.image}[alt="seal"]`);
  await expect(image).toBeVisible();
  await expect(image).toHaveAttribute("src", /^data:image\/png;base64,/);
});

test("a browser paste copies an image from this editor as a new image", async ({
  page,
}) => {
  await openHarness(page, "demo");
  const target = firstTextParagraph(await blocks(page));
  await selectText(page, target.index, 0, Math.min(4, target.docText.length));

  await page.evaluate(
    ({ imageClass, png, sheetClass }) => {
      const data = new DataTransfer();
      data.setData("text/plain", "seal");
      data.setData(
        "text/html",
        `<img class="${imageClass}" src="data:image/png;base64,${png}" ` +
          'alt="copied seal" data-extent=\'{"cx":952500,"cy":952500}\' ' +
          'data-xml="source drawing">'
      );
      const sheet = document.querySelector(`.${sheetClass}`);
      if (!(sheet instanceof HTMLElement)) {
        throw new Error("editor sheet missing");
      }
      sheet.dispatchEvent(
        new ClipboardEvent("paste", {
          bubbles: true,
          cancelable: true,
          clipboardData: data,
        })
      );
    },
    {
      imageClass: editorClassNames.image,
      png: TINY_PNG_BASE64,
      sheetClass: editorClassNames.sheet,
    }
  );

  const image = page.locator(
    `img.${editorClassNames.image}[alt="copied seal"]`
  );
  await expect(image).toBeVisible();
  await expect(image).not.toHaveAttribute("data-xml");
  await expect(image).toHaveAttribute("width", "100");
});
