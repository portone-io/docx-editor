import { expect, type Page, test } from "@playwright/test";
import { editorClassNames } from "../src/styles/classNames";
import {
  blocks,
  openHarness,
  pageScale,
  rightClick,
  selectText,
  settle,
  spaces,
} from "./support/harness";

/**
 * Every place one page parts from the next carries a mark, whether the gap is drawn as a band
 * or the text crosses it, so the pages on the sheet are those marks plus the first page.
 */
const PAGE_BOUNDARIES = `.${editorClassNames.pageSplit}, .${editorClassNames.pageCrossed}`;

async function layoutSnapshot(page: Page) {
  return page.evaluate(
    ({ classes, boundaries }) => {
      const layer = document.querySelector(`.${classes.pageLayer}`);
      if (!(layer instanceof HTMLElement))
        throw new Error("page layer missing");
      return {
        sheetHeight: layer.style.getPropertyValue("--docx-editor-sheet-height"),
        pages: document.querySelectorAll(boundaries).length + 1,
      };
    },
    { classes: editorClassNames, boundaries: PAGE_BOUNDARIES }
  );
}

async function frameSnapshots(page: Page, count: number) {
  return page.evaluate(
    async ({ classes, boundaries, frames }) => {
      const layer = document.querySelector(`.${classes.pageLayer}`);
      if (!(layer instanceof HTMLElement))
        throw new Error("page layer missing");
      const found: string[] = [];
      for (let frame = 0; frame < frames; frame += 1) {
        await new Promise<void>((resolve) =>
          requestAnimationFrame(() => resolve())
        );
        found.push(
          `${layer.style.getPropertyValue("--docx-editor-sheet-height")}/${document.querySelectorAll(boundaries).length + 1}`
        );
      }
      return found;
    },
    { classes: editorClassNames, boundaries: PAGE_BOUNDARIES, frames: count }
  );
}

/**
 * For every footnote drawn beside the paper, the page its area is named for and the page its own
 * reference stands on, read off the screen as it is drawn.
 */
async function footnotePages(page: Page) {
  const ends = await page
    .locator(PAGE_BOUNDARIES)
    .evaluateAll((marks) =>
      marks.map((mark) => mark.getBoundingClientRect().top)
    );
  const pageAt = (y: number) => ends.filter((end) => end <= y).length + 1;
  const areas = await page
    .getByRole("region", { name: /^Footnotes on page \d+$/ })
    .all();
  const found: { area: number; reference: number }[] = [];
  for (const area of areas) {
    const label = (await area.getAttribute("aria-label")) ?? "";
    const number =
      (await area
        .locator(`sup.${editorClassNames.noteMark}`)
        .first()
        .textContent()) ?? "";
    const box = await page
      .locator(`.${editorClassNames.sheet} [aria-label="Footnote ${number}"]`)
      .first()
      .boundingBox();
    if (!box)
      throw new Error(`footnote ${number} has no reference on the sheet`);
    found.push({
      area: Number.parseInt(label.replace(/\D+/g, ""), 10),
      reference: pageAt(box.y),
    });
  }
  return found;
}

/** The zoom the comment panel reads, and the type it sets at it. */
async function commentTypography(page: Page) {
  return page.evaluate((classes) => {
    const workspace = document.querySelector(`.${classes.workspace}`);
    const meta = document.querySelector(`.${classes.commentMeta}`);
    const body = document.querySelector(`.${classes.commentBody}`);
    if (
      !(workspace instanceof HTMLElement) ||
      !(meta instanceof HTMLElement) ||
      !(body instanceof HTMLElement)
    ) {
      throw new Error("comment typography sample missing");
    }
    const styles = getComputedStyle(workspace);
    return {
      zoom: Number.parseFloat(styles.getPropertyValue("--docx-editor-zoom")),
      meta: Number.parseFloat(getComputedStyle(meta).fontSize),
      body: Number.parseFloat(getComputedStyle(body).fontSize),
    };
  }, editorClassNames);
}

/**
 * Where the text stands on the paper, in the paper's own pixels: the scale the paper is drawn at
 * (`pageScale`) is divided back out, so a reading taken at one zoom is comparable with one taken at
 * another. `textBottom` is where the last block of the document ends, which every measurement above
 * it and every space the pagination opened adds into.
 */
async function paperLayout(page: Page, scale: number) {
  return page.evaluate(
    ({ classes, boundaries, scale }) => {
      const layer = document.querySelector(`.${classes.pageLayer}`);
      const sheet = document.querySelector(`.${classes.sheet}`);
      if (!(layer instanceof HTMLElement) || !(sheet instanceof HTMLElement)) {
        throw new Error("page layer missing");
      }
      const sheetTop = sheet.getBoundingClientRect().top;
      const onPaper = (y: number) =>
        Math.round(((y - sheetTop) / scale) * 10) / 10;
      const breaks = Array.from(document.querySelectorAll(boundaries), (mark) =>
        onPaper(mark.getBoundingClientRect().top)
      );
      const opens = new Map<number, number>();
      let textBottom = 0;
      let index = -1;
      for (const block of sheet.children) {
        const box = block.getBoundingClientRect();
        if (box.height === 0) continue;
        index += 1;
        const top = onPaper(box.top);
        // A block standing level with a break opens the page below it
        const standsOn = breaks.filter((at) => at <= top + 1).length + 1;
        if (!opens.has(standsOn)) opens.set(standsOn, index);
        textBottom = onPaper(box.bottom);
      }
      return {
        sheetHeight: layer.style.getPropertyValue("--docx-editor-sheet-height"),
        pages: breaks.length + 1,
        openedByBlock: [...opens.values()],
        textBottom,
      };
    },
    { classes: editorClassNames, boundaries: PAGE_BOUNDARIES, scale }
  );
}

/** What is left to scroll past the foot of the drawn paper, and whether anything scrolls sideways */
async function scrollRoom(page: Page) {
  return page.evaluate((classes) => {
    const scroller = document.querySelector(`.${classes.root}`);
    const layer = document.querySelector(`.${classes.pageLayer}`);
    if (!(scroller instanceof HTMLElement) || !(layer instanceof HTMLElement)) {
      throw new Error("scroll box missing");
    }
    const above = Number.parseFloat(getComputedStyle(scroller).paddingTop);
    return {
      belowThePaper: Math.round(
        scroller.scrollHeight - above - layer.getBoundingClientRect().height
      ),
      sideways: scroller.scrollWidth > scroller.clientWidth,
    };
  }, editorClassNames);
}

/** Whether everything drawn over the paper is drawn within it, said in words a failure can read */
async function marksOnThePaper(page: Page) {
  return page.evaluate(
    ({ classes, boundaries }) => {
      const sheet = document.querySelector(`.${classes.sheet}`);
      if (!(sheet instanceof HTMLElement)) throw new Error("paper missing");
      const paper = sheet.getBoundingClientRect();
      const within = (element: Element) => {
        const box = element.getBoundingClientRect();
        return (
          box.left >= paper.left - 1 &&
          box.right <= paper.right + 1 &&
          box.top >= paper.top - 1 &&
          box.bottom <= paper.bottom + 1
        );
      };
      const say = (found: Element[]) =>
        found.length === 0
          ? "none drawn"
          : found.every(within)
            ? "on the paper"
            : "off the paper";
      return {
        guides: say([...document.querySelectorAll(boundaries)]),
        areas: say([
          ...document.querySelectorAll('[aria-label^="Footnotes on page"]'),
        ]),
      };
    },
    { classes: editorClassNames, boundaries: PAGE_BOUNDARIES }
  );
}

/**
 * The paper is always the width the document names, so a narrow window only scales the sheet it is
 * drawn on. The pages a document breaks into are therefore the document's own, and neither the
 * count nor the sheet height may follow the window.
 *
 * This runs over `notes`, which keeps room at the foot of two of its pages, so what the footnote
 * band asks for is held to the window as well. The document that has almost no room left over is
 * the demo, and the test below holds that one.
 */
test("a document's pages do not follow the width of the window", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1000, height: 900 });
  await openHarness(page, "notes");
  const baseline = await layoutSnapshot(page);
  // A one-page document would hold the count still on its own, so the comparisons below would pass
  // over a pagination that had stopped running at all
  expect(baseline.pages).toBeGreaterThan(1);
  const placement = await footnotePages(page);
  expect(placement).toHaveLength(2);

  for (const width of [839, 719, 559, 419]) {
    await page.setViewportSize({ width, height: 900 });
    await settle(page);
    await settle(page);

    expect(await layoutSnapshot(page)).toEqual(baseline);
    expect(new Set(await frameSnapshots(page, 12)).size).toBe(1);
    // Each footnote keeps to the foot of the page its own reference stands on, whatever the
    // window has scaled the paper to
    const narrow = await footnotePages(page);
    expect(narrow.map((note) => note.area)).toEqual(
      narrow.map((note) => note.reference)
    );
    expect(narrow).toEqual(placement);
  }

  // The zoom the window worked out is not the only one the pages have to survive
  const scale = page.getByLabel("Zoom");
  await scale.selectOption("1");
  await settle(page);
  await settle(page);
  expect(await layoutSnapshot(page)).toEqual(baseline);
});

/**
 * Where a document's pages break is the document's own answer, so it may not follow the scale the
 * paper is being read at, whether the reader chose that scale from the toolbar or the window worked
 * it out for a narrow screen.
 *
 * The demo is the document to hold this over: its fourth page is left with under a pixel of body
 * once the page keeps room at its foot for the footnote its text refers to. While the paper was
 * scaled with the CSS `zoom` property, Chromium laid every box out on whole device pixels of the
 * scaled rendering - a table row came out a third of a pixel taller at 0.75 than at 1, and this
 * document's blocks over a pixel and a half lower by the foot of that page - so below about 0.8 it
 * gained a page holding nothing but an empty paragraph. The layer is scaled by a transform instead
 * (`styles/editor.css`), so the boxes measured are the paper's own whatever it is drawn at.
 *
 * `textBottom` is what carries the teeth: a page count only moves once a boundary is actually
 * crossed, while the foot of the text moves by every fraction of a pixel the measurement drifted.
 */
test("a document's pages break in the same places at every zoom", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await openHarness(page, "demo");
  const zoom = page.getByLabel("Zoom");
  await zoom.selectOption("1");
  await settle(page);
  await settle(page);
  const baseline = await paperLayout(page, await pageScale(page));
  // A one-page document would hold its own answer still, so the comparisons below would pass over a
  // pagination that had stopped running at all
  expect(baseline.pages).toBeGreaterThan(1);
  const opened = await spaces(page);

  for (const factor of ["1.5", "1.25", "0.75", "0.5"]) {
    await zoom.selectOption(factor);
    await settle(page);
    await settle(page);
    // A page break the text carries parts its page from the next by a space of its own, and how
    // tall that space is says where on the paper the break was read
    expect(await spaces(page)).toBe(opened);
    expect(await paperLayout(page, await pageScale(page))).toEqual(baseline);
  }

  // The scale a narrow window works out for itself is the one a reader never asked for
  await zoom.selectOption("fit-width");
  for (const width of [1400, 839, 719, 559, 419]) {
    await page.setViewportSize({ width, height: 900 });
    await settle(page);
    await settle(page);
    expect(await spaces(page)).toBe(opened);
    expect(await paperLayout(page, await pageScale(page))).toEqual(baseline);
  }
});

/**
 * An application is free to scale whatever it mounts the editor inside - a dialog opening on a
 * scale, a container fitting itself to a small screen - and a rectangle read on the paper is then
 * drawn at that scale as well as at the reader's own, so the measurement has to divide by both.
 * Read against the layer's own scale alone, the demo came to 5, 5 and 9 pages under shells of 0.6,
 * 0.85 and 1.4 where it stands at 6.
 */
test("a document's pages ignore a scale the application puts around the editor", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await openHarness(page, "demo");
  const baseline = await paperLayout(page, await pageScale(page));
  expect(baseline.pages).toBeGreaterThan(1);
  const opened = await spaces(page);

  for (const shell of ["scale(0.6)", "scale(0.85)", "scale(1.4)"]) {
    await page.evaluate((transform) => {
      const around = document.body.firstElementChild;
      if (!(around instanceof HTMLElement)) throw new Error("nothing to scale");
      around.style.transformOrigin = "0 0";
      around.style.transform = transform;
    }, shell);
    // A transform moves no box a resize observation would see, so the pages are laid out again by
    // nudging the window: what is read below is a measurement taken while the shell stood
    await page.setViewportSize({ width: 1399, height: 900 });
    await settle(page);
    await page.setViewportSize({ width: 1400, height: 900 });
    await settle(page);
    await settle(page);

    expect(await pageScale(page)).toBeCloseTo(
      Number.parseFloat(shell.replace(/\D*([\d.]+).*/, "$1")),
      5
    );
    expect(await spaces(page)).toBe(opened);
    expect(await paperLayout(page, await pageScale(page))).toEqual(baseline);
  }
});

/**
 * What the reader scrolls through, and what stands over the paper, follow the scale the paper is
 * drawn at rather than the size it was laid out at.
 *
 * A transform leaves the layout box unscaled, so the layer stands outside the flow and the box it
 * stands in holds the room it takes (`ui/usePageRoom`). Were that room the unscaled size, a reader
 * at half scale would scroll through twice the paper, and at one and a half would not reach the
 * end of it.
 */
test("the scroll box and the marks over the paper follow its scale", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await openHarness(page, "notes");
  const zoom = page.getByLabel("Zoom");
  for (const factor of ["1", "1.5", "0.75", "0.5"]) {
    await zoom.selectOption(factor);
    await settle(page);
    await settle(page);

    const room = await scrollRoom(page);
    // The scroll box ends at the foot of the paper plus its own bottom padding, and the paper is
    // narrower than the box at each of these scales, so nothing scrolls sideways
    expect(room).toEqual({ belowThePaper: 24, sideways: false });
    expect(await marksOnThePaper(page)).toEqual({
      guides: "on the paper",
      areas: "on the paper",
    });
  }
});

/**
 * The pages this document breaks into are asserted above, so what is held here is the rail beside
 * them.
 */
test("the demo keeps its comment rail usable in narrow layouts", async ({
  page,
}) => {
  await page.setViewportSize({ width: 1000, height: 900 });
  await openHarness(page, "demo");
  await page.getByRole("button", { name: "Show comments" }).click();
  await settle(page);
  const zoom = page.getByLabel("Zoom");
  await expect(zoom).toHaveValue("fit-width");
  const font = page.getByLabel("Font", { exact: true });
  await expect(font).toHaveCSS("width", "112px");
  await expect(font).toHaveCSS("text-overflow", "ellipsis");
  await expect(font).toHaveCSS("white-space", "nowrap");
  for (const size of [
    {
      width: 839,
      commentWidth: 240,
      canvasPadding: "10px 12px",
      cardPadding: 10,
      metaLadder: 11,
      bodyLadder: 13,
    },
    {
      width: 719,
      commentWidth: 200,
      canvasPadding: "8px",
      cardPadding: 8,
      metaLadder: 10,
      bodyLadder: 11,
    },
    {
      width: 559,
      commentWidth: 180,
      canvasPadding: "6px",
      cardPadding: 6,
      metaLadder: 10,
      bodyLadder: 11,
    },
    {
      width: 419,
      commentWidth: 180,
      canvasPadding: "6px",
      cardPadding: 6,
      metaLadder: 10,
      bodyLadder: 11,
    },
  ]) {
    await page.setViewportSize({ width: size.width, height: 900 });
    await settle(page);
    await settle(page);

    // Whatever this width paginates to, it has to settle on one answer rather than flicker
    expect(new Set(await frameSnapshots(page, 12)).size).toBe(1);
    await expect(page.locator(`.${editorClassNames.commentsPanel}`)).toHaveCSS(
      "width",
      `${size.commentWidth}px`
    );
    await expect(page.locator(`.${editorClassNames.commentsCanvas}`)).toHaveCSS(
      "padding",
      size.canvasPadding
    );
    await expect(
      page.locator(`.${editorClassNames.commentCard}`).first()
    ).toHaveCSS("padding", `${size.cardPadding}px`);
    const editor = page.locator(`.${editorClassNames.root}`);
    await editor.evaluate((element) => {
      element.scrollLeft = element.scrollWidth;
    });
    await settle(page);
    const editorBox = await editor.boundingBox();
    expect(
      await editor.evaluate((element) => {
        if (!(element instanceof HTMLElement)) {
          throw new Error("responsive editor is not an HTML element");
        }
        return element.offsetWidth - element.clientWidth;
      })
    ).toBeGreaterThanOrEqual(10);
    const commentsBox = await page
      .locator(`.${editorClassNames.commentsPanel}`)
      .boundingBox();
    if (!editorBox || !commentsBox) {
      throw new Error("responsive editor layout missing");
    }
    expect(editorBox.x + editorBox.width).toBeLessThanOrEqual(
      commentsBox.x + 1
    );
    const typography = await commentTypography(page);
    expect(await pageScale(page)).toBeCloseTo(typography.zoom, 5);
    expect(typography.meta).toBeCloseTo(
      Math.max(9, size.metaLadder * typography.zoom),
      1
    );
    expect(typography.body).toBeCloseTo(
      Math.max(10, size.bodyLadder * typography.zoom),
      1
    );
  }

  await zoom.selectOption("1");
  await settle(page);
  await settle(page);
  await expect(zoom).toHaveValue("1");
  await expect.poll(() => pageScale(page)).toBe(1);
  expect(new Set(await frameSnapshots(page, 12)).size).toBe(1);
  expect(
    await page
      .locator(`.${editorClassNames.root}`)
      .evaluate((root) => root.scrollWidth > root.clientWidth)
  ).toBe(true);

  await zoom.selectOption("fit-width");
  await settle(page);
  await settle(page);
  await expect.poll(() => pageScale(page)).toBe(0.5);

  const commentHeader = page
    .locator(`.${editorClassNames.commentHeader}`)
    .first();
  await expect(commentHeader).toHaveCSS("flex-direction", "column");
  const author = await commentHeader
    .locator(`.${editorClassNames.commentAuthor}`)
    .boundingBox();
  const date = await commentHeader
    .locator(`.${editorClassNames.commentDate}`)
    .boundingBox();
  const icons = await commentHeader
    .locator(`.${editorClassNames.commentIconActions}`)
    .boundingBox();
  const firstIcon = await commentHeader
    .locator(`.${editorClassNames.commentIconActions} button`)
    .first()
    .boundingBox();
  if (!author || !date || !icons || !firstIcon) {
    throw new Error("responsive comment header missing");
  }
  expect(date.y).toBeGreaterThan(author.y);
  expect(icons.y).toBeGreaterThan(date.y);
  expect(firstIcon.width).toBe(24);
  expect(firstIcon.height).toBe(24);

  await page.getByRole("button", { name: "Text color" }).click();
  const panel = page.locator(`.${editorClassNames.popover}`);
  await expect(panel).toHaveCSS("font-size", "11px");
  const box = await panel.boundingBox();
  if (!box) throw new Error("popover missing");
  expect(box.x).toBeGreaterThanOrEqual(8);
  expect(box.x + box.width).toBeLessThanOrEqual(411);
  expect(box.y).toBeGreaterThanOrEqual(8);
  expect(box.y + box.height).toBeLessThanOrEqual(892);
  expect(
    await panel.evaluate((element) => ({
      horizontal: element.scrollWidth > element.clientWidth,
      vertical: element.scrollHeight > element.clientHeight,
    }))
  ).toEqual({ horizontal: false, vertical: false });
  await page.keyboard.press("Escape");

  const target = (await blocks(page)).find(
    (block) => block.type === "paragraph" && block.docText.length > 6
  );
  if (!target) throw new Error("demo paragraph missing");
  await selectText(page, target.index, 1, 5);
  await rightClick(page);
  const menu = page.getByRole("menu", { name: "Text actions" });
  await expect(menu).toHaveCSS("font-size", "11px");
  const menuBox = await menu.boundingBox();
  const menuIcon = await menu.locator("svg").first().boundingBox();
  if (!menuBox || !menuIcon) throw new Error("text menu missing");
  expect(menuBox.x).toBeGreaterThanOrEqual(8);
  expect(menuBox.x + menuBox.width).toBeLessThanOrEqual(411);
  expect(menuBox.y).toBeGreaterThanOrEqual(8);
  expect(menuBox.y + menuBox.height).toBeLessThanOrEqual(892);
  expect(menuIcon.width).toBe(13);
  expect(menuIcon.height).toBe(13);
});

test("opening and resizing a narrow comment rail preserves left scroll", async ({
  page,
}) => {
  await page.setViewportSize({ width: 419, height: 900 });
  await openHarness(page, "demo");
  const editor = page.locator(`.${editorClassNames.root}`);
  await expect(
    page.locator(`.${editorClassNames.commentsPanel}`)
  ).toHaveAttribute("data-view", "rail");
  await expect
    .poll(() => editor.evaluate((element) => element.scrollLeft))
    .toBe(0);

  await page.setViewportSize({ width: 1000, height: 900 });
  await page.getByLabel("Zoom").selectOption("1");
  await settle(page);
  await settle(page);
  await editor.evaluate((element) => {
    element.scrollLeft = 0;
  });

  await page.setViewportSize({ width: 419, height: 900 });
  await settle(page);
  await settle(page);

  expect(
    await editor.evaluate(
      (element) => element.scrollWidth > element.clientWidth
    )
  ).toBe(true);
  await expect
    .poll(() => editor.evaluate((element) => element.scrollLeft))
    .toBe(0);
});

test("comment typography follows the editor zoom", async ({ page }) => {
  await page.setViewportSize({ width: 1400, height: 900 });
  await openHarness(page, "demo");
  await page.getByRole("button", { name: "Show comments" }).click();
  await settle(page);
  const zoom = page.getByLabel("Zoom");
  for (const factor of [0.5, 0.75, 1, 1.25, 1.5]) {
    await zoom.selectOption(String(factor));
    await settle(page);
    const typography = await commentTypography(page);
    expect(typography.zoom).toBeCloseTo(factor, 5);
    expect(await pageScale(page)).toBeCloseTo(factor, 5);
    // The floor keeps the smaller metas readable while the paper shrinks
    expect(typography.meta).toBeCloseTo(Math.max(9, 12 * factor), 1);
    expect(typography.body).toBeCloseTo(Math.max(10, 14 * factor), 1);
  }
});
