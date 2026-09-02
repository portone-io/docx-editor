// @vitest-environment jsdom
/**
 * The toolbar buttons that style table cells, driven through the real editor.
 *
 * What is checked here is the wiring: the buttons are dead outside a table, a preset reaches
 * the cell holding the caret, and the color a palette picks lands in the exported XML.
 */

import { unzipSync } from "fflate";
import { TextSelection } from "prosemirror-state";
import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { decode, makeDocx } from "../__testing__/docx";
import { EDITING } from "../__testing__/mode";
import { renderInto } from "../__testing__/react";
import { DocxEditor, type DocxEditorHandle } from "../DocxEditor";
import { editorClassNames } from "../styles/classNames";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  host.remove();
});

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

/** A paragraph outside a table, then a table of two cells */
const WITH_TABLE = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>' +
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr>${cellXml("left")}${cellXml("right")}</w:tr>` +
    "</w:tbl>"
);

const WITH_TABLE_BORDERS = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>' +
    "<w:tbl>" +
    "<w:tblPr><w:tblBorders>" +
    '<w:top w:val="single" w:sz="4" w:color="A6B7C8"/>' +
    '<w:left w:val="single" w:sz="4" w:color="A6B7C8"/>' +
    '<w:bottom w:val="single" w:sz="4" w:color="A6B7C8"/>' +
    '<w:right w:val="single" w:sz="4" w:color="A6B7C8"/>' +
    '<w:insideV w:val="single" w:sz="4" w:color="A6B7C8"/>' +
    "</w:tblBorders></w:tblPr>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr>${cellXml("left")}${cellXml("right")}</w:tr>` +
    "</w:tbl>"
);

const WITH_LOCKED_TABLE = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>' +
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
    '<w:tr><w:sdt><w:sdtPr><w:id w:val="7"/>' +
    '<w:lock w:val="sdtContentLocked"/></w:sdtPr>' +
    `<w:sdtContent>${cellXml("locked")}</w:sdtContent></w:sdt></w:tr>` +
    "</w:tbl>"
);

const render = (element: ReactNode) => renderInto(host, element);

function mount(bytes: Uint8Array) {
  const box: { current: DocxEditorHandle | null } = { current: null };
  const unmount = render(
    <DocxEditor
      document={bytes}
      mode={EDITING}
      ref={box}
      renderImportError={() => null}
    />
  );
  const handle = box.current;
  if (!handle) throw new Error("the ref was not attached");
  return { handle, unmount };
}

function button(label: string): HTMLButtonElement {
  const found = host.querySelector(`button[aria-label="${label}"]`);
  if (!(found instanceof HTMLButtonElement)) {
    throw new Error(`button not found: ${label}`);
  }
  return found;
}

function click(label: string): void {
  act(() => button(label).click());
}

function colorIndicator(label: string): HTMLElement {
  const indicator = button(label).querySelector<HTMLElement>(
    `.${editorClassNames.toolbarColorIndicator}`
  );
  if (!indicator) throw new Error(`color indicator not found: ${label}`);
  return indicator;
}

/** A row of the open menu, found by the text written on it */
function clickMenuItem(text: string): void {
  const rows = Array.from(host.querySelectorAll('button[role^="menuitem"]'));
  const row = rows.find((element) => element.textContent === text);
  if (!(row instanceof HTMLButtonElement)) {
    throw new Error(`menu item not found: ${text}`);
  }
  act(() => row.click());
}

/** A swatch of the open palette */
function clickSwatch(hex: string): void {
  const swatch = host.querySelector(`button[aria-label="${hex}"]`);
  if (!(swatch instanceof HTMLButtonElement)) {
    throw new Error(`swatch not found: ${hex}`);
  }
  act(() => swatch.click());
}

/** Puts the caret in the first cell of the table, which is the second body block */
function caretInFirstCell(handle: DocxEditorHandle): void {
  act(() => {
    const { view } = handle;
    const at = view.state.doc.child(0).nodeSize + 4;
    view.dispatch(
      view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(at)))
    );
  });
}

function documentXml(handle: DocxEditorHandle): string {
  return decode(unzipSync(handle.exportBytes())["word/document.xml"]);
}

const CELL_BUTTONS = [
  "Cell vertical alignment",
  "Cell borders",
  "Border color",
  "Cell fill",
];

describe("the cell style buttons", () => {
  it("uses one vertical alignment menu and does not expose cell padding", () => {
    const { handle, unmount } = mount(WITH_TABLE);
    caretInFirstCell(handle);

    expect(host.querySelector('button[aria-label="Cell padding"]')).toBeNull();
    expect(
      button("Cell vertical alignment").querySelector(
        ".lucide-align-vertical-justify-start"
      )
    ).not.toBeNull();

    click("Cell vertical alignment");
    const menu = host.querySelector(
      '[role="menu"][aria-label="Cell vertical alignment"]'
    );
    expect(menu?.querySelectorAll('[role="menuitemradio"]')).toHaveLength(3);
    expect(menu?.querySelector('[aria-checked="true"]')?.textContent).toContain(
      "Align cell top"
    );
    unmount();
  });

  it("are dead while the caret is outside a table", () => {
    const { unmount } = mount(WITH_TABLE);
    for (const label of CELL_BUTTONS) {
      expect(button(label).disabled).toBe(true);
    }
    unmount();
  });

  it("come alive in a cell, except the border color, which waits for a line to color", () => {
    const { handle, unmount } = mount(WITH_TABLE);
    caretInFirstCell(handle);

    expect(button("Cell borders").disabled).toBe(false);
    expect(button("Cell fill").disabled).toBe(false);
    expect(button("Border color").disabled).toBe(true);

    click("Cell borders");
    clickMenuItem("All borders");
    expect(button("Border color").disabled).toBe(false);
    unmount();
  });

  it("edits a visible table-level border without requiring a cell-border preset", () => {
    const { handle, unmount } = mount(WITH_TABLE_BORDERS);
    caretInFirstCell(handle);

    expect(button("Border color").disabled).toBe(false);
    expect(colorIndicator("Border color").style.backgroundColor).toBe(
      "rgb(166, 183, 200)"
    );
    click("Border color");
    clickSwatch("#ff0000");

    const exported = documentXml(handle);
    expect(exported).toContain(
      '<w:top w:val="single" w:sz="4" w:color="A6B7C8"/>'
    );
    expect(exported).toContain(
      '<w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="FF0000"/>'
    );
    unmount();
  });

  it("remain disabled for every style in a content-locked cell", () => {
    const { handle, unmount } = mount(WITH_LOCKED_TABLE);
    caretInFirstCell(handle);

    for (const label of CELL_BUTTONS) {
      expect(button(label).disabled, label).toBe(true);
    }
    unmount();
  });

  it("write the preset, then the color it is drawn in, into the cell", () => {
    const { handle, unmount } = mount(WITH_TABLE);
    caretInFirstCell(handle);

    click("Cell borders");
    clickMenuItem("All borders");
    click("Border color");
    clickSwatch("#ff0000");
    expect(colorIndicator("Border color").style.backgroundColor).toBe(
      "rgb(255, 0, 0)"
    );

    const exported = documentXml(handle);
    expect(exported).toContain('<w:top w:val="single" w:sz="4"');
    expect(exported).toContain('w:color="FF0000"');
    // The cell next to it was never selected, so it keeps its original formatting
    expect(exported).toContain(`${cellXml("right")}`);
    unmount();
  });

  it("fill the cell and take the fill back off again", () => {
    const { handle, unmount } = mount(WITH_TABLE);
    caretInFirstCell(handle);

    expect(colorIndicator("Cell fill").style.backgroundColor).toBe(
      "transparent"
    );
    click("Cell fill");
    clickSwatch("#ffff00");
    expect(colorIndicator("Cell fill").style.backgroundColor).toBe(
      "rgb(255, 255, 0)"
    );
    expect(documentXml(handle)).toContain(
      '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'
    );

    click("Cell fill");
    act(() => {
      const rows = Array.from(host.querySelectorAll("button"));
      const clear = rows.find((element) => element.textContent === "No fill");
      if (!(clear instanceof HTMLButtonElement)) {
        throw new Error("the no-fill entry is missing");
      }
      clear.click();
    });
    expect(documentXml(handle)).not.toContain("w:shd");
    expect(colorIndicator("Cell fill").style.backgroundColor).toBe(
      "transparent"
    );
    unmount();
  });

  it("writes the vertical alignment selected from the menu", () => {
    const { handle, unmount } = mount(WITH_TABLE);
    caretInFirstCell(handle);

    click("Cell vertical alignment");
    clickMenuItem("Align cell bottom");

    const exported = documentXml(handle);
    expect(exported).toContain('<w:vAlign w:val="bottom"/>');
    expect(
      button("Cell vertical alignment").querySelector(
        ".lucide-align-vertical-justify-end"
      )
    ).not.toBeNull();
    unmount();
  });
});
