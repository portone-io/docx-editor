// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { makeFormattedNotesDocx, makeNotesDocx } from "../../__testing__/docx";
import { EDITING } from "../../__testing__/mode";
import { renderInto } from "../../__testing__/react";
import { DocxEditor } from "../../DocxEditor";
import { editorClassNames } from "../../styles/classNames";

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

function found(selector: string): HTMLElement {
  const element = host.querySelector(selector);
  if (!(element instanceof HTMLElement)) {
    throw new Error(`nothing matches ${selector}`);
  }
  return element;
}

function rowTexts(list: HTMLElement): (string | null)[] {
  return [...list.querySelectorAll(`.${editorClassNames.noteRow}`)].map(
    (row) => row.textContent
  );
}

describe("the notes listed after the last page", () => {
  it("lists endnotes with their formatting and labels after the last page", () => {
    const unmount = renderInto(
      host,
      <DocxEditor
        document={makeFormattedNotesDocx()}
        mode={EDITING}
        renderImportError={() => null}
      />
    );

    const sheet = found(`.${editorClassNames.sheet}`);
    const endnotes = found('section[aria-label="Endnotes"]');
    expect(
      sheet.compareDocumentPosition(endnotes) & Node.DOCUMENT_POSITION_FOLLOWING
    ).not.toBe(0);
    // The footnotes stand at the foot of their pages while the pages are drawn
    expect(rowTexts(endnotes)).toEqual(["1 Italic endnote"]);
    expect(
      endnotes.querySelector(`sup.${editorClassNames.noteMark}`)?.textContent
    ).toBe("1");
    const italic = [
      ...endnotes.querySelectorAll<HTMLElement>(`.${editorClassNames.run}`),
    ].find((run) => run.textContent === " Italic endnote");
    expect(italic?.style.fontStyle).toBe("italic");
    unmount();
  });

  it("names a note drawing a mark of its own by its kind alone", () => {
    const unmount = renderInto(
      host,
      <DocxEditor
        document={makeNotesDocx(
          '<w:p><w:r><w:t xml:space="preserve">Marked</w:t></w:r>' +
            '<w:r><w:footnoteReference w:customMarkFollows="1" w:id="2"/></w:r>' +
            "<w:r><w:t>*</w:t></w:r></w:p>"
        )}
        mode={EDITING}
        showPageGuides={false}
        renderImportError={() => null}
      />
    );

    // The reference draws no number, so the row is named the way the reference itself is
    expect(
      [
        ...found('section[aria-label="Footnotes"]').querySelectorAll(
          "[role='group']"
        ),
      ].map((row) => row.getAttribute("aria-label"))
    ).toEqual(["Footnote"]);
    unmount();
  });

  it("lists footnotes ahead of endnotes when page guides are off", () => {
    const unmount = renderInto(
      host,
      <DocxEditor
        document={makeFormattedNotesDocx()}
        mode={EDITING}
        showPageGuides={false}
        renderImportError={() => null}
      />
    );

    expect(
      rowTexts(found('section[aria-label="Footnotes and endnotes"]'))
    ).toEqual([
      "1Plain then bold wordsSecond paragraph",
      "2 Later footnote",
      "1 Italic endnote",
    ]);
    expect(host.querySelector(`.${editorClassNames.footnoteAreas}`)).toBeNull();
    unmount();
  });
});
