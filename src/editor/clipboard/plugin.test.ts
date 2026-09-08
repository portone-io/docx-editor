// @vitest-environment jsdom

import type { Node as PMNode } from "prosemirror-model";
import { Plugin, TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";
import { beforeEach, describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeStyledDocx,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { toRunFormat } from "../../model/format";
import { editorClassNames } from "../../styles/classNames";
import { createEditorView, editorStateForSession } from "../createEditor";
import { defineClipboardEvent } from "./__testing__/clipboardEvent";
import { docxClipboard } from "./plugin";

const HEADING_STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
  '<w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="DocumentHeading1">' +
  '<w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>';

const PARAGRAPH =
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>';

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

const TABLE =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${cellXml("A1")}${cellXml("B1")}</w:tr>` +
  `<w:tr>${cellXml("A2")}${cellXml("B2")}</w:tr>` +
  "</w:tbl>";

function openEditor(
  bytes: Uint8Array,
  plugins: readonly Plugin[] = []
): { view: EditorView; session: SessionStore } {
  const { doc, session } = importDocx(bytes);
  const view = createEditorView({
    mount: document.createElement("div"),
    state: editorStateForSession(
      { doc, session },
      { consumerPlugins: plugins }
    ),
    onStateChange: () => {},
  });
  return { view, session };
}

/** A paste the way the browser delivers one, since jsdom has no clipboard of its own */
function paste(view: EditorView, data: Record<string, string>): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => data[type] ?? "" },
  });
  view.dom.dispatchEvent(event);
}

/** Holding Shift while pasting, which is the keypress ProseMirror reads the plain paste off */
function pressShift(view: EditorView): void {
  view.dom.dispatchEvent(
    new KeyboardEvent("keydown", { key: "V", shiftKey: true, bubbles: true })
  );
}

function cellTexts(doc: PMNode): string[] {
  const texts: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "tableCell") texts.push(node.textContent);
    return true;
  });
  return texts;
}

function selectAllCells(view: EditorView): void {
  const cells: number[] = [];
  view.state.doc.descendants((node, pos) => {
    if (node.type.name === "tableCell") cells.push(pos);
    return true;
  });
  const first = cells[0];
  const last = cells[cells.length - 1];
  if (first === undefined || last === undefined) {
    throw new Error("the document holds no table");
  }
  view.dispatch(
    view.state.tr.setSelection(
      CellSelection.create(view.state.doc, first, last)
    )
  );
}

beforeEach(() => {
  defineClipboardEvent();
});

describe("the clipboard plugin", () => {
  it("routes view.pasteHTML through the same reader as a paste event", () => {
    const html = "<h1>Heading</h1>";
    const byEvent = openEditor(makeStyledDocx(PARAGRAPH, HEADING_STYLES));
    paste(byEvent.view, { "text/html": html, "text/plain": "Heading" });
    const byCall = openEditor(makeStyledDocx(PARAGRAPH, HEADING_STYLES));
    byCall.view.pasteHTML(html);

    expect(documentXmlOf(byCall.view.state.doc, byCall.session)).toBe(
      documentXmlOf(byEvent.view.state.doc, byEvent.session)
    );
    expect(documentXmlOf(byCall.view.state.doc, byCall.session)).toContain(
      '<w:pStyle w:val="DocumentHeading1"/>'
    );
    byEvent.view.destroy();
    byCall.view.destroy();
  });

  it("never runs a schema parseDOM rule on clipboard DOM", () => {
    const { view } = openEditor(makeDocx(PARAGRAPH));
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
    );
    paste(view, {
      "text/plain": "smuggled",
      "text/html":
        `<p class="${editorClassNames.paragraph}" data-pm-slice="0 0 []" ` +
        'data-ppr="&lt;w:pPr&gt;&lt;w:jc w:val=&quot;center&quot;/&gt;&lt;/w:pPr&gt;" ' +
        'data-pattrs="w:rsidR=&quot;00AB12&quot;">smuggled</p>',
    });

    const paragraph = view.state.doc.firstChild;
    expect(paragraph?.textContent).toBe("smuggled");
    // Both would have come over as they stand had the schema's own rule read this markup
    expect(paragraph?.attrs.pPr).toBeNull();
    expect(paragraph?.attrs.pAttrs).toBeNull();
    view.destroy();
  });

  it("closes a foreign block paste so a heading lands as its own paragraph", () => {
    const { view } = openEditor(makeStyledDocx(PARAGRAPH, HEADING_STYLES));
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 4))
    );
    paste(view, { "text/plain": "Heading", "text/html": "<h1>Heading</h1>" });

    expect(
      view.state.doc.children.map((paragraph) => paragraph.textContent)
    ).toEqual(["sou", "Heading", "rce"]);
    view.destroy();
  });

  it("pastes a paragraph over a cell selection into every selected cell", () => {
    const { view } = openEditor(makeDocx(TABLE));
    selectAllCells(view);
    paste(view, { "text/plain": "x", "text/html": "<p>x</p>" });

    expect(cellTexts(view.state.doc)).toEqual(["x", "x", "x", "x"]);
    view.destroy();
  });

  it("keeps a Shift paste as plain text with w:br line breaks", () => {
    const { view, session } = openEditor(makeDocx(PARAGRAPH));
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
    );
    pressShift(view);
    paste(view, {
      "text/plain": "a\nb",
      "text/html": '<p><span style="font-weight: 700">a<br>b</span></p>',
    });

    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.textContent).toBe("ab");
    expect(documentXmlOf(view.state.doc, session)).toContain("<w:br/>");
    let bold = false;
    view.state.doc.descendants((node) => {
      for (const mark of node.marks) {
        if (toRunFormat(mark.attrs.format)?.bold === true) bold = true;
      }
      return true;
    });
    expect(bold).toBe(false);
    view.destroy();
  });

  it("dispatches the paste with the uiEvent meta", () => {
    const events: unknown[] = [];
    const watcher = new Plugin({
      appendTransaction(transactions) {
        for (const tr of transactions) {
          if (tr.docChanged) events.push(tr.getMeta("uiEvent"));
        }
        return null;
      },
    });
    const { view } = openEditor(makeDocx(PARAGRAPH), [watcher]);
    paste(view, { "text/plain": "pasted", "text/html": "<p>pasted</p>" });

    // A `handlePaste` of its own would put the content in itself, and the meta ProseMirror's own
    // insertion carries - which a consumer plugin reads a paste off - would never be written
    expect(docxClipboard().props.handlePaste).toBeUndefined();
    expect(events).toEqual(["paste"]);
    view.destroy();
  });
});
