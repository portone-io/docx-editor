import { expect, type Page, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  blocks,
  docText,
  firstTextParagraph,
  openHarness,
  selectText,
  settle,
} from "./support/harness";

const TINY_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNg" +
  "AAIAAAUAAen63NgAAAAASUVORK5CYII=";

async function browserPaste(
  page: Page,
  html: string,
  text: string
): Promise<void> {
  await page.evaluate(
    ({ html, text, sheetClass }) => {
      const data = new DataTransfer();
      if (html) data.setData("text/html", html);
      if (text) data.setData("text/plain", text);
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
    },
    { html, text, sheetClass: editorClassNames.sheet }
  );
}

test("Shift paste ignores an HTML image and the next ordinary paste loads it", async ({
  page,
}) => {
  let requests = 0;
  await page.route("https://assets.example/plain-paste.png", (route) => {
    requests += 1;
    return route.fulfill({
      body: Buffer.from(TINY_PNG_BASE64, "base64"),
      contentType: "image/png",
      headers: { "access-control-allow-origin": "*" },
    });
  });
  await openHarness(page, "demo");
  const target = firstTextParagraph(await blocks(page));
  await selectText(page, target.index, 0, Math.min(4, target.docText.length));
  const html =
    '<p><b>Formatted text</b><img src="https://assets.example/plain-paste.png" alt="plain paste image"></p>';
  await page.keyboard.down("Shift");
  await browserPaste(page, html, "Text only");
  await page.keyboard.up("Shift");
  expect(await docText(page)).toContain("Text only");
  const image = page.locator(
    `img.${editorClassNames.image}[alt="plain paste image"]`
  );
  await expect(image).toHaveCount(0);
  expect(requests).toBe(0);

  await browserPaste(page, html, "Text only");
  await expect(image).toBeVisible();
  expect(requests).toBe(1);
});

test("unreadable HTML uses its text fallback and empty text does not delete the selection", async ({
  page,
}) => {
  await openHarness(page, "demo");
  const target = firstTextParagraph(await blocks(page));
  await selectText(page, target.index, 0, Math.min(4, target.docText.length));
  await browserPaste(page, "<!-- producer metadata -->", "Fallback text");
  expect(await docText(page)).toContain("Fallback text");
  await selectText(page, target.index, 0, 4);
  const before = await docText(page);
  await browserPaste(page, "", "\u0001");
  expect(await docText(page)).toBe(before);
});

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

  // A real copy out of this editor carries no `data-extent`: the document's own measure stays
  // behind and the size travels as the pixels the picture was drawn at
  await browserPaste(
    page,
    `<img class="${editorClassNames.image}" ` +
      `src="data:image/png;base64,${TINY_PNG_BASE64}" ` +
      'alt="drawn seal" width="150" height="75">',
    "seal"
  );
  const drawn = page.locator(`img.${editorClassNames.image}[alt="drawn seal"]`);
  await expect(drawn).toBeVisible();
  await expect(drawn).toHaveAttribute("width", "150");
  await expect(drawn).toHaveAttribute("height", "75");
});
