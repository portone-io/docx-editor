// @vitest-environment jsdom
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import {
  makeControlSectionHeadersFootersDocx,
  makeHeadersFootersDocx,
  makeTwoSectionHeadersFootersDocx,
} from "../__testing__/docx";
import { type HeadersFooters, variantsFor } from "../docx/headersFooters";
import { importDocx } from "../docx/importDocx";
import { sectionsOf } from "../docx/sections";
import { storyNodeOf } from "../schema/stories";
import { editorClassNames } from "../styles/classNames";
import { PageGuides } from "./PageGuides";
import { pageLayout, sectionPixels } from "./pageLayout";
import type { PageFace, PageOverlay } from "./usePageLayout";

/** One page of a document written in a single section, where the two counts are the same number */
function face(page: number, top: number): PageFace {
  return {
    page,
    section: 0,
    pageInSection: page,
    headerTop: top,
    footerTop: top + 920,
    left: 80,
    width: 640,
    crossed: false,
    reserved: [],
    trailing: null,
  };
}

const overlay: PageOverlay = {
  left: 0,
  top: 0,
  width: 800,
  sheetHeight: 3000,
  marks: [],
  pages: [face(1, 40), face(2, 1040), face(3, 2040)],
};

/** The stories the document's one section shows, which is what a page of it draws */
function openedStories(): HeadersFooters {
  const { doc, session } = importDocx(makeHeadersFootersDocx());
  const section = sectionsOf(doc)[0];
  if (!section) throw new Error("the test document has no section");
  return variantsFor(section, session.headerFooterStories, (key) =>
    storyNodeOf(doc, key)
  );
}

let host: HTMLDivElement | null = null;
let root: Root | null = null;

afterEach(() => {
  if (root) act(() => root?.unmount());
  root = null;
  host?.remove();
  host = null;
});

function render(element: React.ReactElement): HTMLDivElement {
  host = document.createElement("div");
  document.body.appendChild(host);
  root = createRoot(host);
  const live = root;
  act(() => live.render(element));
  return host;
}

/** What every page of the rendered overlay draws as its header */
function headers(drawn: HTMLDivElement): (string | null)[] {
  return Array.from(
    drawn.querySelectorAll(`.${editorClassNames.pageHeader}`)
  ).map((element) => element.textContent);
}

describe("page header and footer guides", () => {
  it("draws the selected story and evaluated fields on every visual page", () => {
    const headersFooters = openedStories();
    const drawn = render(
      <PageGuides overlay={overlay} headersFootersFor={() => headersFooters} />
    );

    expect(headers(drawn)).toEqual([
      "First header",
      "Default 5 of 3",
      "Even header",
    ]);
    expect(
      Array.from(drawn.querySelectorAll(`.${editorClassNames.pageFooter}`)).map(
        (element) => element.textContent
      )
    ).toEqual(["First footer", "Default footer", "Even footer"]);
  });

  it("draws each page with the stories of the section it belongs to", () => {
    const headersFooters = openedStories();
    const evenOnly: HeadersFooters = {
      ...headersFooters,
      firstPageDifferent: false,
      headers: {
        ...headersFooters.headers,
        default: headersFooters.headers.even,
      },
    };
    const drawn = render(
      <PageGuides
        overlay={{ ...overlay, pages: overlay.pages.slice(0, 2) }}
        headersFootersFor={(page) =>
          page.page === 1 ? headersFooters : evenOnly
        }
      />
    );

    expect(headers(drawn)).toEqual(["First header", "Even header"]);
  });

  it("asks the section each page opens in", () => {
    const { doc, session } = importDocx(makeTwoSectionHeadersFootersDocx());
    const shown = sectionsOf(doc).map((section) =>
      variantsFor(section, session.headerFooterStories, (key) =>
        storyNodeOf(doc, key)
      )
    );
    // The second paragraph opens the body's section, which is the one that declares w:titlePg
    const secondSection: PageFace = {
      ...face(2, 1040),
      section: 1,
      pageInSection: 1,
    };
    const drawn = render(
      <PageGuides
        overlay={{ ...overlay, pages: [face(1, 40), secondSection] }}
        headersFootersFor={(page) => shown[page.section] ?? null}
      />
    );

    // The first section draws its own default story, and the second draws the first-page story
    // of its own first page even though that page is the document's second
    expect(headers(drawn)).toEqual(["Even header", "First header"]);
  });

  it("draws the next section's header on a page opened inside a control", () => {
    const { doc, session } = importDocx(makeControlSectionHeadersFootersDocx());
    const sections = sectionsOf(doc);
    const shown = sections.map((section) =>
      variantsFor(section, session.headerFooterStories, (key) =>
        storyNodeOf(doc, key)
      )
    );
    const control = doc.child(0);
    const second = 1 + (control.firstChild?.nodeSize ?? 0);
    // Both paragraphs stand in one block, parted where the first section ends
    const layout = pageLayout({
      blocks: [
        {
          pos: 0,
          gap: 0,
          height: 200,
          breakBefore: false,
          breakAfter: false,
          candidates: [
            { at: second, offset: 100, forced: false, repeatHeight: 0 },
          ],
          minFirstPiece: 100,
          keepWithNext: false,
        },
      ],
      sections: sectionPixels(sections),
    });
    const drawn = render(
      <PageGuides
        overlay={{
          ...overlay,
          pages: layout.pages.map((start) => ({
            ...face(start.page, 40 + (start.page - 1) * 1000),
            section: start.section,
            pageInSection: start.pageInSection,
          })),
        }}
        headersFootersFor={(page) => shown[page.section] ?? null}
      />
    );

    expect(control.type.name).toBe("sdtBlock");
    expect(layout.pages.map((page) => page.section)).toEqual([0, 1]);
    // The page opening with the control's second paragraph is the body section's first page
    expect(headers(drawn)).toEqual(["Even header", "First header"]);
  });

  it("projects a story's direct paragraph alignment", () => {
    const headersFooters = openedStories();
    const right = headersFooters.headers.default;
    if (!right) throw new Error("the test document has no default header");
    const drawn = render(
      <PageGuides
        overlay={{ ...overlay, pages: [face(1, 40)] }}
        headersFootersFor={() => ({
          ...headersFooters,
          headers: {
            default: { ...right, align: "right" },
            first: { ...right, align: "right" },
            even: { ...right, align: "right" },
          },
        })}
      />
    );

    expect(
      drawn.querySelector<HTMLElement>(`.${editorClassNames.pageHeader}`)?.style
        .textAlign
    ).toBe("right");
  });
});
