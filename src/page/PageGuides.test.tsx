// @vitest-environment jsdom
import { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it } from "vitest";
import { makeHeadersFootersDocx } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { editorClassNames } from "../styles/classNames";
import { PageGuides } from "./PageGuides";
import type { PageOverlay } from "./usePageLayout";

const overlay: PageOverlay = {
  left: 0,
  top: 0,
  width: 800,
  sheetHeight: 3000,
  marks: [],
  badges: [],
  pages: [
    {
      page: 1,
      headerTop: 40,
      footerTop: 960,
      left: 80,
      width: 640,
      crossed: false,
    },
    {
      page: 2,
      headerTop: 1040,
      footerTop: 1960,
      left: 80,
      width: 640,
      crossed: false,
    },
    {
      page: 3,
      headerTop: 2040,
      footerTop: 2960,
      left: 80,
      width: 640,
      crossed: false,
    },
  ],
};

let host: HTMLDivElement | null = null;

afterEach(() => {
  host?.remove();
  host = null;
});

describe("page header and footer guides", () => {
  it("draws the selected story and evaluated fields on every visual page", () => {
    host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);
    const { headersFooters } = importDocx(makeHeadersFootersDocx()).session;

    act(() =>
      root.render(
        <PageGuides overlay={overlay} headersFooters={headersFooters} />
      )
    );

    expect(
      Array.from(host.querySelectorAll(`.${editorClassNames.pageHeader}`)).map(
        (element) => element.textContent
      )
    ).toEqual(["First header", "Default 5 of 3", "Even header"]);
    expect(
      Array.from(host.querySelectorAll(`.${editorClassNames.pageFooter}`)).map(
        (element) => element.textContent
      )
    ).toEqual(["First footer", "Default footer", "Even footer"]);

    act(() => root.unmount());
  });

  it("projects a story's direct paragraph alignment", () => {
    host = document.createElement("div");
    document.body.appendChild(host);
    const root = createRoot(host);
    const headersFooters = importDocx(makeHeadersFootersDocx()).session
      .headersFooters;
    const right = {
      segments: [{ kind: "text" as const, value: "Header" }],
      align: "right" as const,
    };

    act(() =>
      root.render(
        <PageGuides
          overlay={{ ...overlay, pages: [overlay.pages[0]!] }}
          headersFooters={{
            ...headersFooters,
            headers: { default: right, first: right, even: right },
          }}
        />
      )
    );

    expect(
      host.querySelector<HTMLElement>(`.${editorClassNames.pageHeader}`)?.style
        .textAlign
    ).toBe("right");
    act(() => root.unmount());
  });
});
