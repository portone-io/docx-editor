// @vitest-environment jsdom
import { undo } from "prosemirror-history";
import { DOMSerializer, type Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeStyledDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { toParagraphFormat, toRunFormat } from "../../model/format";
import { docxSchema } from "../../schema";
import { addRowAfter } from "../../table";
import { createEditorState } from "../createEditor";
import { insertTable } from "../insertTable";

/** A default paragraph style that lays down both paragraph and character formatting */
const NORMAL_STYLE =
  '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
  '<w:name w:val="Normal"/>' +
  '<w:pPr><w:jc w:val="center"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>';

const HEADING_STYLE =
  '<w:style w:type="paragraph" w:styleId="Heading1">' +
  '<w:name w:val="heading 1"/>' +
  '<w:rPr><w:i/><w:sz w:val="40"/></w:rPr></w:style>';

const BODY = '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>';

/** The state built out of everything the document told us, the same as `DocxEditor` builds it */
function opened(styles = NORMAL_STYLE): {
  state: EditorState;
  session: SessionStore;
} {
  const { doc, session } = importDocx(makeStyledDocx(BODY, styles));
  return {
    state: createEditorState(doc, {
      styles: session.styles,
      defaults: session.defaults,
      paragraphDefaults: session.paragraphDefaults,
      paragraphStyles: session.paragraphStyles,
    }),
    session,
  };
}

function ran(state: EditorState, command: Command): EditorState {
  let next = state;
  expect(
    command(state, (tr) => {
      next = state.apply(tr);
    })
  ).toBe(true);
  return next;
}

function undone(state: EditorState): EditorState {
  let next = state;
  undo(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

/** The paragraph inside the cell at this spot of the first table, and the values it draws with */
function cellParagraph(doc: PMNode, row: number, col: number) {
  const table = doc.child(1);
  const cell = table.child(row).child(col);
  const paragraph = cell.child(0);
  return {
    node: paragraph,
    format: toParagraphFormat(paragraph.attrs.format),
    styleRun: toRunFormat(paragraph.attrs.styleRun),
  };
}

/** Types into the first cell of the table, where inserting one leaves the caret */
function typedInFirstCell(state: EditorState, text: string): EditorState {
  return state.apply(state.tr.insertText(text, state.selection.head));
}

/** The paragraph as it is drawn on screen */
function rendered(paragraph: PMNode): string {
  const host = document.createElement("div");
  host.appendChild(
    DOMSerializer.fromSchema(docxSchema).serializeNode(paragraph)
  );
  return host.innerHTML;
}

/** A paragraph carrying no style values, which is the one shape this plugin writes to */
function freshHeading(text = "heading"): PMNode {
  return docxSchema.nodes.paragraph.create(
    { pPr: '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' },
    [docxSchema.text(text)]
  );
}

let view: EditorView | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
});

function mounted(state: EditorState): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  view = new EditorView(mount, { state });
  return view;
}

/** The composition events `view.composing` follows, which is all this plugin reads */
function composition(live: EditorView, open: boolean): void {
  live.dom.dispatchEvent(
    new CompositionEvent(open ? "compositionstart" : "compositionend", {
      bubbles: true,
      data: "",
    })
  );
}

/** Puts a fresh paragraph in with the caret inside it, which is where a composition would sit */
function putHeadingUnderTheCaret(live: EditorView): void {
  const at = live.state.doc.child(0).nodeSize;
  const tr = live.state.tr.insert(at, freshHeading());
  live.dispatch(tr.setSelection(TextSelection.create(tr.doc, at + 2)));
}

function headingStyleRun(live: EditorView) {
  return toRunFormat(live.state.doc.child(1).attrs.styleRun);
}

describe("a paragraph an edit built from nothing", () => {
  it("draws text typed into a new table's cell in the default style", () => {
    const inserted = ran(opened().state, insertTable({ rows: 2, columns: 2 }));
    const typed = typedInFirstCell(inserted, "cell");
    const cell = cellParagraph(typed.doc, 0, 0);

    expect(cell.node.textContent).toBe("cell");
    // The text carries no run of its own, so the paragraph is what has to draw it
    expect(cell.node.child(0).marks).toEqual([]);
    expect(cell.styleRun).toEqual({ bold: true, fontSizePt: 11 });
    expect(cell.format).toEqual({ align: "center" });
    expect(rendered(cell.node)).toContain("font-weight: bold");
  });

  it("draws it the same in a cell an added row brought", () => {
    const inserted = ran(opened().state, insertTable({ rows: 2, columns: 2 }));
    const grown = ran(inserted, addRowAfter);
    const typed = typedInFirstCell(grown, "new cell");
    const cell = cellParagraph(typed.doc, 1, 0);

    expect(cell.node.textContent).toBe("new cell");
    expect(cell.styleRun).toEqual({ bold: true, fontSizePt: 11 });
  });

  // The values go in as an appended transaction, which the history takes as part of the same event
  it("is taken back together with the edit that built it, in one undo", () => {
    const state = opened().state;
    const inserted = ran(state, insertTable({ rows: 2, columns: 2 }));
    expect(cellParagraph(inserted.doc, 0, 0).styleRun).toEqual({
      bold: true,
      fontSizePt: 11,
    });

    const back = undone(inserted);
    expect(back.doc.childCount).toBe(1);
    expect(back.doc.child(0).textContent).toBe("body");
  });

  it("is read under the style it names itself", () => {
    const { state } = opened(NORMAL_STYLE + HEADING_STYLE);
    const put = state.apply(
      state.tr.insert(state.doc.child(0).nodeSize, freshHeading())
    );

    expect(toRunFormat(put.doc.child(1).attrs.styleRun)).toEqual({
      italic: true,
      fontSizePt: 20,
    });
  });

  /**
   * An untouched block goes back out as its original XML, and that rests on the node still being
   * the same one, so a paragraph already carrying its values may never be written again
   */
  it("leaves a paragraph that already carries its values as the very same node", () => {
    const inserted = ran(opened().state, insertTable({ rows: 2, columns: 2 }));
    const body = inserted.doc.child(0);
    const typed = typedInFirstCell(inserted, "cell");
    const untouchedCell = cellParagraph(typed.doc, 1, 1).node;

    // The paragraph the document came with is untouched by the cells being filled in
    expect(typed.doc.child(0)).toBe(body);

    const again = typed.apply(
      typed.tr.insertText("more", typed.selection.head)
    );
    expect(again.doc.child(0)).toBe(body);
    expect(cellParagraph(again.doc, 1, 1).node).toBe(untouchedCell);
  });
});

/**
 * Rewriting a node redraws it, and a redraw under an open composition is what the browser answers
 * by taking the composition down. So the paragraph an IME is composing in waits.
 */
describe("a paragraph an IME is composing in", () => {
  it("is written to as usual while nothing is being composed", () => {
    const live = mounted(opened(NORMAL_STYLE + HEADING_STYLE).state);
    putHeadingUnderTheCaret(live);

    expect(headingStyleRun(live)).toEqual({ italic: true, fontSizePt: 20 });
  });

  it("is left as it stands until the composition is over", () => {
    const live = mounted(opened(NORMAL_STYLE + HEADING_STYLE).state);
    composition(live, true);
    expect(live.composing).toBe(true);

    putHeadingUnderTheCaret(live);
    expect(headingStyleRun(live)).toBeNull();

    // Every further edit under the same composition leaves it alone as well
    live.dispatch(live.state.tr.insertText("more", live.state.selection.head));
    expect(headingStyleRun(live)).toBeNull();

    // And the next edit after the composition is the one that reads the styles into it
    composition(live, false);
    expect(live.composing).toBe(false);
    live.dispatch(live.state.tr.insertText("!", 1));

    expect(headingStyleRun(live)).toEqual({ italic: true, fontSizePt: 20 });
  });

  it("does not hold up the other paragraphs of the document", () => {
    const live = mounted(opened(NORMAL_STYLE + HEADING_STYLE).state);
    composition(live, true);
    putHeadingUnderTheCaret(live);

    // A second fresh paragraph, this one nowhere near the composition
    const tr = live.state.tr.insert(
      live.state.doc.content.size,
      freshHeading()
    );
    live.dispatch(tr);

    expect(headingStyleRun(live)).toBeNull();
    expect(toRunFormat(live.state.doc.child(2).attrs.styleRun)).toEqual({
      italic: true,
      fontSizePt: 20,
    });
  });
});
