// @vitest-environment jsdom
import { act } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { makeNotesDocx } from "../__testing__/docx";
import { renderInto } from "../__testing__/react";
import { DocxEditor } from "../DocxEditor";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => host.remove());

describe("the document notes panel", () => {
  it("shows superscript references and the plain-text note bodies", () => {
    let unmount = () => {};
    act(() => {
      unmount = renderInto(
        host,
        <DocxEditor document={makeNotesDocx()} renderImportError={() => null} />
      );
    });

    expect(host.querySelector('[aria-label="Footnote 1"]')?.textContent).toBe(
      "1"
    );
    expect(host.querySelector('[aria-label="Endnote 1"]')?.textContent).toBe(
      "1"
    );
    const panel = host.querySelector('section[aria-label="Document notes"]');
    expect(
      panel?.parentElement?.classList.contains("docx-editor-page-layer")
    ).toBe(true);
    expect(panel?.querySelector("h2")?.textContent).toBe(
      "Footnotes and endnotes"
    );
    expect((panel as HTMLElement | null)?.style.width).not.toBe("");
    expect(panel?.textContent).toContain("Footnote 1Footnote body");
    expect(panel?.textContent).toContain("Second line");
    expect(panel?.textContent).toContain("Endnote 1Endnote body");
    unmount();
  });
});
