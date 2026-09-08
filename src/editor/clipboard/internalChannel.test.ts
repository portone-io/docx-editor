// @vitest-environment jsdom

import type { Node as PMNode } from "prosemirror-model";
import { NodeSelection, TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";
import { beforeEach, describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeNumberedDocx,
  makeStyledDocx,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { toRunFormat } from "../../model/format";
import { newListsOf } from "../../numbering/listRegistry";
import { templateList } from "../../numbering/listTemplate";
import { docxSchema } from "../../schema";
import {
  listRefOf,
  removeFromList,
  toggleBulletList,
} from "../commands/listCommands";
import { createEditorView, editorStateForSession } from "../createEditor";
import { defineClipboardEvent } from "./__testing__/clipboardEvent";

const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
  '<w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/>' +
  "<w:rPr><w:i/></w:rPr></w:style>";

/** Everything the reproduction of 06 §2.5 puts in one paragraph */
const LOADED_PARAGRAPH =
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
  '<w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:b/><w:caps/>' +
  '<w:lang w:val="fr-FR"/></w:rPr><w:t>Caps</w:t></w:r>' +
  "<w:r><w:tab/></w:r>" +
  '<w:r><w:br w:type="page"/></w:r></w:p>';

const TAIL = "<w:p><w:r><w:t>tail</w:t></w:r></w:p>";

const SOURCE_PARAGRAPH =
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>';

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

const table = (a: string, b: string, c: string, d: string) =>
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${cellXml(a)}${cellXml(b)}</w:tr>` +
  `<w:tr>${cellXml(c)}${cellXml(d)}</w:tr>` +
  "</w:tbl>";

interface Opened {
  view: EditorView;
  session: SessionStore;
}

function open(bytes: Uint8Array): Opened {
  const { doc, session } = importDocx(bytes);
  const view = createEditorView({
    mount: document.createElement("div"),
    state: editorStateForSession({ doc, session }),
    onStateChange: () => {},
  });
  return { view, session };
}

/** What a copy leaves on the clipboard, as the browser would hand it to a paste */
function copy(view: EditorView): Record<string, string> {
  const { dom, text } = view.serializeForClipboard(
    view.state.selection.content()
  );
  return { "text/html": dom.innerHTML, "text/plain": text };
}

function paste(view: EditorView, data: Record<string, string>): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => data[type] ?? "" },
  });
  view.dom.dispatchEvent(event);
}

/**
 * Fakes the dragstart ProseMirror writes the dragged slice on. jsdom has neither DragEvent nor
 * DataTransfer, and it measures no text, so no position can be found from the drag's coordinates.
 */
function dragOut(view: EditorView): void {
  const event = new Event("dragstart", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "dataTransfer", {
    value: {
      files: [],
      clearData: () => {},
      setData: () => {},
      effectAllowed: "",
    },
  });
  view.posAtCoords = () => null;
  view.dom.dispatchEvent(event);
}

function selectBlock(view: EditorView, pos: number): void {
  view.dispatch(
    view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos))
  );
}

function caretAtEnd(view: EditorView): void {
  view.dispatch(
    view.state.tr.setSelection(
      TextSelection.create(view.state.doc, view.state.doc.content.size - 1)
    )
  );
}

function occurrences(xml: string, fragment: string): number {
  return xml.split(fragment).length - 1;
}

function cellTexts(view: EditorView): string[] {
  const texts: string[] = [];
  view.state.doc.descendants((node) => {
    if (node.type === docxSchema.nodes.tableCell) texts.push(node.textContent);
    return true;
  });
  return texts;
}

function selectCells(view: EditorView, from: number, to: number): void {
  const cells: number[] = [];
  view.state.doc.descendants((node, pos) => {
    if (node.type === docxSchema.nodes.tableCell) cells.push(pos);
    return true;
  });
  const anchor = cells[from];
  const head = cells[to];
  if (anchor === undefined || head === undefined) {
    throw new Error("the document holds too few cells");
  }
  view.dispatch(
    view.state.tr.setSelection(
      CellSelection.create(view.state.doc, anchor, head)
    )
  );
}

beforeEach(() => {
  defineClipboardEvent();
});

describe("the internal clipboard channel", () => {
  it("pastes a copied paragraph back with its pPr, rPr, tabs and page break intact", () => {
    const { view, session } = open(
      makeStyledDocx(LOADED_PARAGRAPH + TAIL, STYLES)
    );
    selectBlock(view, 0);
    const copied = copy(view);
    // Nothing the document is written with leaves the editor; the token stands for it instead
    expect(copied["text/html"]).not.toMatch(/data-(rpr|ppr|pattrs|xml)=/);
    expect(copied["text/html"]).not.toContain("<w:");

    caretAtEnd(view);
    paste(view, copied);

    const xml = documentXmlOf(view.state.doc, session);
    expect(view.state.doc.childCount).toBe(3);
    expect(occurrences(xml, '<w:jc w:val="center"/>')).toBe(2);
    expect(occurrences(xml, "<w:caps/>")).toBe(2);
    expect(occurrences(xml, '<w:rStyle w:val="Emphasis"/>')).toBe(2);
    expect(occurrences(xml, '<w:lang w:val="fr-FR"/>')).toBe(2);
    expect(occurrences(xml, "<w:tab/>")).toBe(2);
    expect(occurrences(xml, '<w:br w:type="page"/>')).toBe(2);
    view.destroy();
  });

  it("writes the pasted copy from the model rather than from the block it came from", () => {
    const { view, session } = open(
      makeStyledDocx(LOADED_PARAGRAPH + TAIL, STYLES)
    );
    selectBlock(view, 0);
    const copied = copy(view);
    caretAtEnd(view);
    paste(view, copied);

    expect(view.state.doc.lastChild?.attrs.srcId).toBeNull();
    expect(view.state.doc.firstChild?.attrs.srcId).not.toBeNull();
    // One paragraph is the fragment the file arrived as, the other is written from the model
    expect(
      occurrences(
        documentXmlOf(view.state.doc, session),
        '<w:jc w:val="center"/>'
      )
    ).toBe(2);
    view.destroy();
  });

  it("pastes a copied table back as a table", () => {
    const { view } = open(makeDocx(table("A1", "B1", "A2", "B2") + TAIL));
    selectBlock(view, 0);
    const copied = copy(view);
    caretAtEnd(view);
    paste(view, copied);

    const tables = view.state.doc.children.filter(
      (child) => child.type === docxSchema.nodes.table
    );
    expect(tables).toHaveLength(2);
    expect(cellTexts(view)).toEqual([
      "A1",
      "B1",
      "A2",
      "B2",
      "A1",
      "B1",
      "A2",
      "B2",
    ]);
    view.destroy();
  });

  it("keeps a copied link as a link when it is pasted back", () => {
    const { view, session } = open(
      makeDocx(
        '<w:p><w:hyperlink w:anchor="top"><w:r><w:t>jump</w:t></w:r>' +
          "</w:hyperlink></w:p>" +
          TAIL
      )
    );
    selectBlock(view, 0);
    const copied = copy(view);
    caretAtEnd(view);
    paste(view, copied);

    expect(
      occurrences(
        documentXmlOf(view.state.doc, session),
        '<w:hyperlink w:anchor="top">'
      )
    ).toBe(2);
    view.destroy();
  });

  it("lets tableEditing paste a copied cell grid over a cell selection", () => {
    const { view } = open(
      makeDocx(table("A1", "B1", "A2", "B2") + table("x", "x", "x", "x"))
    );
    selectCells(view, 0, 3);
    const copied = copy(view);
    selectCells(view, 4, 7);
    paste(view, copied);

    expect(cellTexts(view)).toEqual([
      "A1",
      "B1",
      "A2",
      "B2",
      "A1",
      "B1",
      "A2",
      "B2",
    ]);
    view.destroy();
  });

  it("falls back to the HTML reader when the token is unknown", () => {
    const { view } = open(makeStyledDocx(LOADED_PARAGRAPH + TAIL, STYLES));
    selectBlock(view, 0);
    const copied = copy(view);
    const forged = {
      ...copied,
      "text/html": copied["text/html"]?.replace(
        /data-docx-clip="[^"]*"/,
        'data-docx-clip="invented"'
      ) as string,
    };
    caretAtEnd(view);
    paste(view, forged);

    // The reader that read the markup keeps none of what only the slice carried
    expect(view.state.doc.lastChild?.attrs.pPr).toBeNull();
    expect(view.state.doc.lastChild?.textContent).toContain("Caps");
    view.destroy();
  });

  it("does not trust the token of a copy a later one replaced", () => {
    const { view } = open(makeStyledDocx(LOADED_PARAGRAPH + TAIL, STYLES));
    selectBlock(view, 0);
    const stale = copy(view);
    selectBlock(view, view.state.doc.firstChild?.nodeSize ?? 0);
    copy(view);
    caretAtEnd(view);
    paste(view, stale);

    // The markup it arrived with is what is read, not whichever slice is being kept
    expect(view.state.doc.lastChild?.textContent).toContain("Caps");
    expect(view.state.doc.lastChild?.attrs.pPr).toBeNull();
    view.destroy();
  });

  it("leaves the clipboard copy standing when a drag is begun and let go", () => {
    const { view } = open(makeStyledDocx(LOADED_PARAGRAPH + TAIL, STYLES));
    selectBlock(view, 0);
    const copied = copy(view);
    // A dragstart writes the dragged slice through the very path a copy is written on
    selectBlock(view, view.state.doc.firstChild?.nodeSize ?? 0);
    dragOut(view);

    caretAtEnd(view);
    paste(view, copied);

    // What the clipboard still holds comes back through the channel rather than through the reader
    expect(view.state.doc.lastChild?.attrs.pPr).toContain(
      '<w:jc w:val="center"/>'
    );
    expect(view.state.doc.lastChild?.textContent).toContain("Caps");
    view.destroy();
  });

  it("sends a copy from another editor's session through the HTML reader", () => {
    const source = open(makeStyledDocx(LOADED_PARAGRAPH + TAIL, STYLES));
    selectBlock(source.view, 0);
    const copied = copy(source.view);
    expect(copied["text/html"]).toContain("data-docx-clip=");

    const destination = open(makeStyledDocx(TAIL, STYLES));
    caretAtEnd(destination.view);
    paste(destination.view, copied);

    const pasted = destination.view.state.doc.lastChild;
    expect(pasted?.attrs.pPr).toBeNull();
    expect(pasted?.attrs.srcId).toBeNull();
    // Only what the markup states survives: the drawn weight, not the character style behind it
    const rPr: unknown = pasted?.firstChild?.marks[0]?.attrs.rPr;
    expect(toRunFormat(pasted?.firstChild?.marks[0]?.attrs.format)?.bold).toBe(
      true
    );
    expect(rPr).not.toContain("rStyle");
    expect(rPr).not.toContain("w:lang");
    expect(rPr).not.toContain("w:caps");
    source.view.destroy();
    destination.view.destroy();
  });

  it("carries the token and nothing else the document is written with", () => {
    const { view } = open(makeStyledDocx(LOADED_PARAGRAPH + TAIL, STYLES));
    selectBlock(view, 0);
    const html = copy(view)["text/html"] ?? "";

    expect(html).toMatch(/data-docx-clip="[^"]+"/);
    expect(
      html.match(/\sdata-(?!style=|pm-slice=|docx-clip=)[\w-]+/g)
    ).toBeNull();
    view.destroy();
  });

  it("drops a bookmark the copied range carried", () => {
    const { view, session } = open(
      makeDocx(
        '<w:p><w:bookmarkStart w:id="1" w:name="mark"/>' +
          "<w:r><w:t>anchored</w:t></w:r>" +
          '<w:bookmarkEnd w:id="1"/></w:p>' +
          TAIL
      )
    );
    selectBlock(view, 0);
    const copied = copy(view);
    caretAtEnd(view);
    paste(view, copied);

    const xml = documentXmlOf(view.state.doc, session);
    expect(view.state.doc.lastChild?.textContent).toBe("anchored");
    expect(occurrences(xml, "<w:bookmarkStart")).toBe(1);
    expect(occurrences(xml, "<w:bookmarkEnd")).toBe(1);
    view.destroy();
  });

  it("brings a copied bulleted list back bulleted once nothing defines its number", () => {
    const { view } = open(makeNumberedDocx(SOURCE_PARAGRAPH + TAIL));
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
    );
    expect(toggleBulletList(view.state, view.dispatch)).toBe(true);
    const listed = view.state.doc.firstChild;
    expect(listRefOf(listed as PMNode)?.numId).toBe(2);

    selectBlock(view, 0);
    const copied = copy(view);
    // Leaving the list gives its number back, so nothing defines it by the time the copy lands
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
    );
    expect(removeFromList(view.state, view.dispatch)).toBe(true);
    expect(newListsOf(view.state.doc.attrs.newLists).size).toBe(0);

    caretAtEnd(view);
    paste(view, copied);

    // The number is issued again; what it stands for is what it stood for where it was copied
    expect(newListsOf(view.state.doc.attrs.newLists).get(2)).toEqual(
      templateList("bullet")
    );
    expect(listRefOf(view.state.doc.lastChild as PMNode)).toEqual({
      numId: 2,
      ilvl: 0,
    });
    view.destroy();
  });

  it("keeps a Shift paste on the plain text path", () => {
    const { view } = open(makeStyledDocx(LOADED_PARAGRAPH + TAIL, STYLES));
    selectBlock(view, 0);
    const copied = copy(view);
    caretAtEnd(view);
    view.dom.dispatchEvent(
      new KeyboardEvent("keydown", { key: "V", shiftKey: true, bubbles: true })
    );
    paste(view, copied);

    expect(view.state.doc.childCount).toBe(2);
    expect(view.state.doc.lastChild?.textContent).toContain("Caps");
    view.destroy();
  });
});
