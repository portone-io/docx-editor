// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import {
  type EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { docxSchema } from "../../schema";
import { unlockAllowed } from "../../schema/locks";
import { deleteColumn, deleteRow, mergeCells } from "../../table";
import { buildResizeColumnTransaction } from "../../table/resize";
import { canRunCommand } from "../commands/canRunCommand";
import { setParagraphAlign } from "../commands/paragraphCommands";
import { createEditorState } from "../createEditor";
import { documentGeometry } from "../documentStyles";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const lockedPr =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';
const openPr = '<w:sdtPr><w:id w:val="7"/></w:sdtPr>';

const sdt = (inner: string, pr: string) =>
  `<w:sdt>${pr}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** One paragraph: plain text, a control, then plain text again */
function paragraph(pr: string): string {
  return `<w:p>${run("a")}${sdt(run("bc"), pr)}${run("d")}</w:p>`;
}

function opened(body: string): EditorState {
  return createEditorState(importDocx(makeDocx(body)).doc);
}

/** The position just before the first character of this text */
function posOf(doc: PMNode, needle: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found < 0 && node.isText && node.text === needle) found = pos;
  });
  if (found < 0) throw new Error(`text not found: ${needle}`);
  return found;
}

/** Whether the transaction went through: a refused one leaves the document as it was */
function applies(state: EditorState, tr: Transaction): boolean {
  return !state.apply(tr).doc.eq(state.doc);
}

function withSelection(
  state: EditorState,
  from: number,
  to: number
): EditorState {
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, from, to))
  );
}

describe("editing inside a locked content control", () => {
  const LOCKED = paragraph(lockedPr);

  it("refuses a character typed inside the control", () => {
    const state = opened(LOCKED);
    const inside = posOf(state.doc, "bc") + 1;
    expect(applies(state, state.tr.insertText("x", inside))).toBe(false);
  });

  it("refuses a deletion inside the control", () => {
    const state = opened(LOCKED);
    const inside = posOf(state.doc, "bc") + 1;
    expect(applies(state, state.tr.delete(inside, inside + 1))).toBe(false);
  });

  /**
   * The bold command leaves a locked stretch out of the range it marks all by itself, so what is
   * asked here is the guard underneath it: a mark laid straight across the locked text
   */
  it("refuses formatting the locked text", () => {
    const state = opened(LOCKED);
    const from = posOf(state.doc, "bc");
    const bold = docxSchema.marks.run.create({ rPr: "<w:rPr><w:b/></w:rPr>" });

    expect(applies(state, state.tr.addMark(from, from + 2, bold))).toBe(false);
  });

  it("refuses a wider deletion that swallows the control", () => {
    const state = opened(LOCKED);
    const paragraphEnd = state.doc.child(0).content.size + 1;
    expect(applies(state, state.tr.delete(1, paragraphEnd))).toBe(false);
  });

  it("accepts a character typed against either edge of the control", () => {
    const state = opened(LOCKED);
    const start = posOf(state.doc, "bc");
    expect(applies(state, state.tr.insertText("x", start))).toBe(true);
    expect(applies(state, state.tr.insertText("x", start + 2))).toBe(true);
  });

  it("leaves the text around the control editable", () => {
    const state = opened(LOCKED);
    expect(applies(state, state.tr.insertText("x", 1))).toBe(true);
    expect(applies(state, state.tr.delete(1, 2))).toBe(true);
  });

  /** The control sits inside the paragraph, so the paragraph's own formatting is not its business */
  it("still lets the paragraph holding the control be aligned", () => {
    const state = opened(LOCKED);
    const selected = withSelection(state, 1, 2);

    let next = selected;
    setParagraphAlign("center")(selected, (tr) => {
      next = selected.apply(tr);
    });
    expect(next.doc.child(0).attrs.pPr).toContain('<w:jc w:val="center"/>');
  });

  it("lets a transaction carrying the unlocking pass through", () => {
    const state = opened(LOCKED);
    const inside = posOf(state.doc, "bc") + 1;
    const tr = state.tr.setMeta(unlockAllowed, true).delete(inside, inside + 1);
    expect(applies(state, tr)).toBe(true);
  });

  it("leaves a control that does not lock its contents freely editable", () => {
    const state = opened(paragraph(openPr));
    const inside = posOf(state.doc, "bc") + 1;
    expect(applies(state, state.tr.insertText("x", inside))).toBe(true);
    expect(applies(state, state.tr.delete(inside, inside + 1))).toBe(true);
  });
});

describe("content put in wearing a lock", () => {
  const LOCKED = paragraph(lockedPr);

  /** What dragging a stretch of text with Alt held does: the slice goes in and nothing is deleted */
  it("is refused wherever the slice lands", () => {
    const state = opened(LOCKED);
    const from = posOf(state.doc, "bc");
    const copied = state.doc.slice(from, from + 2);

    expect(applies(state, state.tr.replace(1, 1, copied))).toBe(false);
    expect(applies(state, state.tr.replace(0, 0, copied))).toBe(false);
  });

  it("is accepted where the slice carries no lock", () => {
    const state = opened(LOCKED);
    const copied = state.doc.slice(1, 2);
    expect(applies(state, state.tr.replace(1, 1, copied))).toBe(true);
  });
});

const cell = (inner: string) => `<w:tc><w:p>${inner}</w:p></w:tc>`;

/** A two by two table whose top left cell stands inside a control carrying these properties */
const cellTable = (pr: string) =>
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${sdt(cell(run("TopLeft")), pr)}${cell(run("TopRight"))}</w:tr>` +
  `<w:tr>${cell(run("BottomLeft"))}${cell(run("BottomRight"))}</w:tr>` +
  "</w:tbl>";

/** That table with the control that shuts the cell outright */
const LOCKED_CELL_TABLE = cellTable(lockedPr);

/** The position of the cell holding this text */
function cellPos(doc: PMNode, needle: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (
      found < 0 &&
      node.type.name === "tableCell" &&
      node.textContent === needle
    ) {
      found = pos;
    }
    return true;
  });
  if (found < 0) throw new Error(`no cell holding: ${needle}`);
  return found;
}

/** The state with the caret inside the cell holding this text */
function inCell(state: EditorState, needle: string): EditorState {
  const at = cellPos(state.doc, needle) + 2;
  return withSelection(state, at, at);
}

describe("editing inside a locked table cell", () => {
  it("refuses a character typed in it and a deletion out of it", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const inside = posOf(state.doc, "TopLeft") + 1;
    expect(applies(state, state.tr.insertText("x", inside))).toBe(false);
    expect(applies(state, state.tr.delete(inside, inside + 1))).toBe(false);
  });

  it("refuses formatting the paragraph the cell holds", () => {
    const state = inCell(opened(LOCKED_CELL_TABLE), "TopLeft");
    let next = state;
    setParagraphAlign("center")(state, (tr) => {
      next = state.apply(tr);
    });
    expect(next.doc.eq(state.doc)).toBe(true);
  });

  it("refuses an edit that would take the cell away with it", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const locked = inCell(state, "TopLeft");
    expect(canRunCommand(deleteRow, locked)).toBe(false);
    expect(canRunCommand(deleteColumn, locked)).toBe(false);

    const merged = state.apply(
      state.tr.setSelection(
        CellSelection.create(
          state.doc,
          cellPos(state.doc, "TopLeft"),
          cellPos(state.doc, "TopRight")
        )
      )
    );
    expect(canRunCommand(mergeCells, merged)).toBe(false);
  });

  it("leaves the rows and columns the lock does not stand in alone", () => {
    const free = inCell(opened(LOCKED_CELL_TABLE), "BottomRight");
    expect(canRunCommand(deleteRow, free)).toBe(true);
    expect(canRunCommand(deleteColumn, free)).toBe(true);
  });

  /** The width is what the cell records about itself rather than what it holds */
  it("leaves the column the cell stands in free to be resized", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const resize = buildResizeColumnTransaction(
      state,
      { tablePos: 0, col: 0, width: 1400 },
      documentGeometry(state)
    );
    if (!resize) throw new Error("the resize could not be built");
    expect(applies(state, resize)).toBe(true);
  });

  it("leaves the cells beside it editable", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const beside = posOf(state.doc, "TopRight") + 1;
    expect(applies(state, state.tr.insertText("x", beside))).toBe(true);
  });

  it("refuses a copy of the table, which would plant the lock a second time", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const copied = state.doc.slice(0, state.doc.content.size);
    expect(applies(state, state.tr.replace(0, 0, copied))).toBe(false);
  });

  it("lets a transaction carrying the unlocking pass through", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const inside = posOf(state.doc, "TopLeft") + 1;
    const tr = state.tr.setMeta(unlockAllowed, true).delete(inside, inside + 1);
    expect(applies(state, tr)).toBe(true);
  });

  /**
   * The lock is written in an attribute of the cell, and the steps that write a cell's width are
   * free to run, so the lock has to be shut against those very steps: otherwise anything that can
   * resize a column could open a locked cell on the way.
   */
  it("refuses a step that puts the lock down by itself", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const pos = cellPos(state.doc, "TopLeft");
    const cellNode = state.doc.nodeAt(pos);
    if (!cellNode) throw new Error("no cell");

    expect(
      applies(state, state.tr.setNodeAttribute(pos, "sdtContentsLocked", false))
    ).toBe(false);
    expect(
      applies(state, state.tr.setNodeAttribute(pos, "sdtDeletionLocked", false))
    ).toBe(false);
    expect(
      applies(
        state,
        state.tr.setNodeMarkup(pos, null, {
          ...cellNode.attrs,
          sdtContentsLocked: false,
          sdtDeletionLocked: false,
        })
      )
    ).toBe(false);
  });

  it("takes the same step once it carries the unlocking", () => {
    const state = opened(LOCKED_CELL_TABLE);
    const pos = cellPos(state.doc, "TopLeft");
    const tr = state.tr
      .setMeta(unlockAllowed, true)
      .setNodeAttribute(pos, "sdtContentsLocked", false);
    expect(applies(state, tr)).toBe(true);
  });
});

const lockPr = (val: string) =>
  `<w:sdtPr><w:id w:val="7"/><w:lock w:val="${val}"/></w:sdtPr>`;

/**
 * The four values `w:lock` takes and the two independent clauses each of them settles
 * (`spec/notes/contentControls`): whether the contents may be edited, and whether the control may
 * be deleted in its entirety.
 */
const LOCK_VALUES = [
  { val: "unlocked", editable: true, deletable: true },
  { val: "sdtLocked", editable: true, deletable: false },
  { val: "contentLocked", editable: false, deletable: true },
  { val: "sdtContentLocked", editable: false, deletable: false },
] as const;

describe.each(LOCK_VALUES)(
  "a control whose lock reads $val",
  ({ val, editable, deletable }) => {
    const BODY = paragraph(lockPr(val));

    it(`takes an edit to its contents: ${editable}`, () => {
      const state = opened(BODY);
      const inside = posOf(state.doc, "bc") + 1;
      expect(applies(state, state.tr.insertText("x", inside))).toBe(editable);
      expect(applies(state, state.tr.delete(inside, inside + 1))).toBe(
        editable
      );
    });

    /** A mark laid across the whole control leaves it standing, so it is an edit of its contents */
    it(`takes formatting laid across the whole of it: ${editable}`, () => {
      const state = opened(BODY);
      const from = posOf(state.doc, "bc");
      const bold = docxSchema.marks.run.create({
        rPr: "<w:rPr><w:b/></w:rPr>",
      });
      expect(applies(state, state.tr.addMark(from, from + 2, bold))).toBe(
        editable
      );
    });

    it(`takes a deletion covering the control exactly: ${deletable}`, () => {
      const state = opened(BODY);
      const from = posOf(state.doc, "bc");
      expect(applies(state, state.tr.delete(from, from + 2))).toBe(deletable);
    });

    /** The paragraph's whole text goes, which takes the control with it and nothing less than it */
    it(`takes a wider deletion that swallows the control: ${deletable}`, () => {
      const state = opened(BODY);
      const paragraphEnd = state.doc.child(0).content.size + 1;
      expect(applies(state, state.tr.delete(1, paragraphEnd))).toBe(deletable);
    });
  }
);

describe.each(LOCK_VALUES)(
  "a cell inside a control whose lock reads $val",
  ({ val, editable, deletable }) => {
    const BODY = cellTable(lockPr(val));

    it(`takes an edit to what the cell holds: ${editable}`, () => {
      const state = opened(BODY);
      const inside = posOf(state.doc, "TopLeft") + 1;
      expect(applies(state, state.tr.insertText("x", inside))).toBe(editable);
      expect(applies(state, state.tr.delete(inside, inside + 1))).toBe(
        editable
      );
    });

    /** Both of these cover the cell whole, which is the control being taken away as one */
    it(`takes the deletion of its row and of its column: ${deletable}`, () => {
      const at = inCell(opened(BODY), "TopLeft");
      expect(canRunCommand(deleteRow, at)).toBe(deletable);
      expect(canRunCommand(deleteColumn, at)).toBe(deletable);
    });

    /** The width is what the cell records about itself rather than what it holds */
    it("leaves the column the cell stands in free to be resized", () => {
      const state = opened(BODY);
      const resize = buildResizeColumnTransaction(
        state,
        { tablePos: 0, col: 0, width: 1400 },
        documentGeometry(state)
      );
      if (!resize) throw new Error("the resize could not be built");
      expect(applies(state, resize)).toBe(true);
    });
  }
);

/** A line break is a node a mark can be put on, unlike the text around it */
const breakRun = "<w:r><w:br/></w:r>";

function posOfBreak(doc: PMNode): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found < 0 && node.type.name === "hardBreak") found = pos;
  });
  if (found < 0) throw new Error("no line break");
  return found;
}

describe("a mark put on a node itself", () => {
  const marked = (state: EditorState, at: number) =>
    state.tr.addNodeMark(
      at,
      docxSchema.marks.run.create({ rPr: "<w:rPr><w:b/></w:rPr>" })
    );

  it("is refused on a node standing inside a locked control", () => {
    const state = opened(`<w:p>${run("a")}${sdt(breakRun, lockedPr)}</w:p>`);
    expect(applies(state, marked(state, posOfBreak(state.doc)))).toBe(false);
  });

  it("is accepted on a node standing outside one", () => {
    const state = opened(`<w:p>${run("a")}${breakRun}</w:p>`);
    expect(applies(state, marked(state, posOfBreak(state.doc)))).toBe(true);
  });
});

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

/** The composition events `view.composing` follows */
function composition(live: EditorView, open: boolean): void {
  live.dom.dispatchEvent(
    new CompositionEvent(open ? "compositionstart" : "compositionend", {
      bubbles: true,
      data: "",
    })
  );
}

/** The refusal is answered on the frame after it */
function nextFrame(): Promise<void> {
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
  });
}

/**
 * The browser takes a composition down whose text was taken back out of the DOM under it and sends
 * no `compositionend` for it, so a refusal has to end the composition itself
 */
describe("a refusal under an open composition", () => {
  it("ends the composition, and the document stands as it was", async () => {
    const live = mounted(opened(paragraph(lockedPr)));
    composition(live, true);
    expect(live.composing).toBe(true);

    const before = live.state.doc;
    const inside = posOf(live.state.doc, "bc") + 1;
    live.dispatch(live.state.tr.insertText("x", inside));
    expect(live.state.doc).toBe(before);

    await nextFrame();
    expect(live.composing).toBe(false);
    expect(live.state.doc).toBe(before);
  });

  it("leaves a composition nothing refused alone", async () => {
    const live = mounted(opened(paragraph(openPr)));
    composition(live, true);

    const inside = posOf(live.state.doc, "bc") + 1;
    live.dispatch(live.state.tr.insertText("x", inside));
    expect(live.state.doc.textContent).toBe("abxcd");

    await nextFrame();
    expect(live.composing).toBe(true);
  });

  it("does nothing where no composition is open", async () => {
    const live = mounted(opened(paragraph(lockedPr)));
    const before = live.state;
    const inside = posOf(live.state.doc, "bc") + 1;
    live.dispatch(live.state.tr.insertText("x", inside));

    await nextFrame();
    expect(live.state).toBe(before);
  });
});
