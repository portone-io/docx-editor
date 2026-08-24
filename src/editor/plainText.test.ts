// @vitest-environment jsdom
import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";
import { describe, expect, it } from "vitest";
import { decode, documentXmlOf, makeDocx } from "../__testing__/docx";
import { exportDocx } from "../docx/exportDocx";
import { NO_DOCUMENT_DEFAULTS } from "../docx/formatting";
import { importDocx } from "../docx/importDocx";
import type { SessionStore } from "../docx/session";
import { documentHasLocked } from "./commands/lockCommands";
import { createEditorState, createEditorView } from "./createEditor";
import { insertPlainText, insertPlainTextAt } from "./plainText";

function openEditor(readOnly = false): {
  view: EditorView;
  session: SessionStore;
} {
  const { doc, session } = importDocx(
    makeDocx('<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>')
  );
  const view = createEditorView({
    mount: document.createElement("div"),
    state: createEditorState(doc),
    defaults: session.defaults,
    readOnly,
    onStateChange: () => {},
  });
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, 2))
  );
  return { view, session };
}

describe("inserting plain text", () => {
  it("a line break becomes a break element instead of splitting the paragraph", () => {
    const { view, session } = openEditor();
    insertPlainText(view, "a\r\nb");

    expect(view.state.doc.childCount).toBe(1);
    const documentXml = decode(
      unzipSync(exportDocx(view.state.doc, session))["word/document.xml"]
    );
    expect(documentXml).toContain("sa");
    expect(documentXml).toContain("<w:br/>");
    expect(documentXml).toContain("bource");
    view.destroy();
  });

  it("a tab becomes a tab element", () => {
    const { view, session } = openEditor();
    insertPlainText(view, "a\tb");

    expect(documentXmlOf(view.state.doc, session)).toContain("<w:tab/>");
    let marked = false;
    view.state.doc.descendants((node) => {
      if (
        node.isText &&
        node.text === "\t" &&
        node.marks.some((mark) => mark.type.name === "tab")
      ) {
        marked = true;
      }
    });
    expect(marked).toBe(true);
    view.destroy();
  });

  it("the vertical tabs Word writes to the clipboard also become line breaks", () => {
    const { view, session } = openEditor();
    insertPlainText(view, "a\u000Bb\u000Cc");

    expect(view.state.doc.childCount).toBe(1);
    expect(
      documentXmlOf(view.state.doc, session).match(/<w:br\/>/g)
    ).toHaveLength(2);
    view.destroy();
  });

  /**
   * The text stays Korean here so that the stripping is covered over multibyte UTF-8.
   * The Word clipboard sometimes carries meaningless control characters mixed in.
   * Inserting them as they are would leave the exported file invalid as XML, so Word could not
   * open it.
   */
  it("strips meaningless control characters so the file that comes out can be opened again", () => {
    const { view, session } = openEditor();
    insertPlainText(view, "가\u0007나\u001F다");

    const bytes = exportDocx(view.state.doc, session);
    const documentXml = decode(unzipSync(bytes)["word/document.xml"]);
    expect(documentXml).toContain("가나다");
    // biome-ignore lint/suspicious/noControlCharactersInRegex: this assertion checks that no control characters remain
    expect(documentXml).not.toMatch(/[\u0000-\u0008\u000E-\u001F]/);
    expect(() => importDocx(bytes)).not.toThrow();
    view.destroy();
  });

  it("changes nothing for empty text", () => {
    const { view } = openEditor();
    const before = view.state.doc;
    insertPlainText(view, "");
    expect(view.state.doc).toBe(before);
    view.destroy();
  });

  it("can replace a mapped document range for an asynchronous fallback", () => {
    const { view, session } = openEditor();
    insertPlainTextAt(
      view,
      "fallback\tline\nnext",
      0,
      view.state.doc.content.size
    );

    expect(view.state.doc.textContent).toBe("fallback\tlinenext");
    const xml = documentXmlOf(view.state.doc, session);
    expect(xml).toContain("<w:tab/>");
    expect(xml).toContain("<w:br/>");
    view.destroy();
  });

  it("is not editable when readOnly", () => {
    const { view } = openEditor(true);
    expect(view.editable).toBe(false);
    view.destroy();
  });
});

const LOCKED_PR =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

const runXml = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A one by two table whose left cell stands inside a control that shuts it */
const LOCKED_CELL_TABLE =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  "<w:tr>" +
  `<w:sdt>${LOCKED_PR}<w:sdtContent>` +
  `<w:tc><w:p>${runXml("locked cell")}</w:p></w:tc>` +
  "</w:sdtContent></w:sdt>" +
  `<w:tc><w:p>${runXml("next cell")}</w:p></w:tc>` +
  "</w:tr></w:tbl>";

/** The position just before each cell, which is what a cell selection is built from */
function cellPositions(doc: PMNode): number[] {
  const found: number[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === "tableCell") found.push(pos);
    return true;
  });
  return found;
}

/**
 * A paste the way the browser delivers one, since jsdom has no clipboard of its own.
 * Both the HTML and the plain text a real copy leaves behind are handed over, so the editor
 * chooses between them for itself.
 */
function paste(view: EditorView, data: Record<string, string>): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => data[type] ?? "" },
  });
  view.dom.dispatchEvent(event);
}

describe("pasting what was copied out of a locked cell", () => {
  it("brings the characters over and leaves the lock behind", () => {
    const { doc } = importDocx(makeDocx(LOCKED_CELL_TABLE));
    const view = createEditorView({
      mount: document.createElement("div"),
      state: createEditorState(doc),
      defaults: NO_DOCUMENT_DEFAULTS,
      readOnly: false,
      onStateChange: () => {},
    });

    const cells = cellPositions(view.state.doc);
    view.dispatch(
      view.state.tr.setSelection(CellSelection.create(view.state.doc, cells[0]))
    );
    const copied = view.serializeForClipboard(view.state.selection.content());
    // The clipboard really does carry the lock, so dropping it is the paste path's own doing
    expect(copied.dom.innerHTML).toContain('data-sdt-contents-locked="1"');

    // The caret at the end of the neighbouring cell, where the copy is pasted
    const target = cellPositions(view.state.doc)[1];
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.near(view.state.doc.resolve(target + 1), 1)
      )
    );
    paste(view, {
      "text/plain": copied.text,
      "text/html": copied.dom.innerHTML,
    });

    const pasted = view.state.doc.nodeAt(cellPositions(view.state.doc)[1]);
    if (!pasted) throw new Error("the cell pasted into is gone");
    expect(pasted.textContent).toBe("locked cellnext cell");
    // Neither the cell's own lock nor a control around the text came along with it
    expect(pasted.attrs.sdtContentsLocked).toBe(false);
    expect(pasted.attrs.sdtPrefix).toBeNull();
    expect(documentHasLocked(pasted)).toBe(false);
    view.destroy();
  });
});
