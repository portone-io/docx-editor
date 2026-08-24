// @vitest-environment jsdom
import { type EditorState, TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { documentXmlOf, makeDocx } from "../__testing__/docx";
import { posOfText } from "../__testing__/editing";
import { importDocx } from "../docx/importDocx";
import { createEditorState } from "./createEditor";
import { docxKeymap } from "./plugins/keymap";

function withCaretAt(state: EditorState, at: number): EditorState {
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, at))
  );
}

/** Places the cursor and presses a single shortcut */
function press(state: EditorState, key: string, at: number): EditorState {
  const withCaret = withCaretAt(state, at);
  let next = withCaret;
  const handled = docxKeymap[key](withCaret, (tr) => {
    next = withCaret.apply(tr);
  });
  expect(handled).toBe(true);
  return next;
}

/** Whether the shortcut has anything to do at this position */
function handles(state: EditorState, key: string, at: number): boolean {
  return docxKeymap[key](withCaretAt(state, at));
}

describe("text editing", () => {
  it("inserted text follows the formatting of the run at that spot", () => {
    const rPr = '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>';
    const bytes = makeDocx(
      `<w:p><w:r>${rPr}<w:t xml:space="preserve">source</w:t></w:r></w:p>`
    );
    const { doc, session } = importDocx(bytes);
    const state = createEditorState(doc);
    const edited = state.apply(state.tr.insertText("added", 2));

    const documentXml = documentXmlOf(edited.doc, session);
    expect(documentXml).toContain(
      `<w:r>${rPr}<w:t xml:space="preserve">saddedource`
    );
    // No formatting absent from the original gets baked in
    expect(documentXml.match(/<w:rPr>/g)).toHaveLength(1);
  });
});

describe("moving between cells with Tab", () => {
  const cell = (text: string) =>
    `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
  const tableBody =
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr>${cell("Left")}${cell("Right")}</w:tr>` +
    "</w:tbl>";

  function tableState(): EditorState {
    const { doc } = importDocx(makeDocx(tableBody));
    return createEditorState(doc);
  }

  it("Tab moves to the cell on the right and Shift+Tab to the cell on the left", () => {
    const state = tableState();
    const forward = press(state, "Tab", posOfText(state.doc, "Left"));
    expect(forward.selection.$head.parent.textContent).toBe("Right");

    const back = press(forward, "Shift-Tab", posOfText(forward.doc, "Right"));
    expect(back.selection.$head.parent.textContent).toBe("Left");
  });

  it("inserts a document tab outside a table", () => {
    const { doc } = importDocx(
      makeDocx('<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>')
    );
    const state = createEditorState(doc);

    const inserted = press(state, "Tab", posOfText(doc, "body"));
    expect(inserted.doc.child(0).textContent).toBe("b\tody");
    expect(handles(state, "Shift-Tab", posOfText(doc, "body"))).toBe(false);
  });
});

describe("splitting a paragraph with Enter", () => {
  it("the new paragraph inherits the original paragraph's formatting", () => {
    const pPr = '<w:pPr><w:jc w:val="center"/></w:pPr>';
    const bytes = makeDocx(
      `<w:p>${pPr}<w:r><w:t xml:space="preserve">frontback</w:t></w:r></w:p>`
    );
    const { doc, session } = importDocx(bytes);
    const split = press(createEditorState(doc), "Enter", 6);

    expect(split.doc.child(0).textContent).toBe("front");
    expect(split.doc.child(1).textContent).toBe("back");
    expect(split.doc.child(1).attrs.format).toEqual({ align: "center" });

    const documentXml = documentXmlOf(split.doc, session);
    expect(documentXml).toContain(
      `<w:p>${pPr}<w:r><w:t xml:space="preserve">front</w:t></w:r></w:p>` +
        `<w:p>${pPr}<w:r><w:t xml:space="preserve">back</w:t></w:r></w:p>`
    );
  });
});
