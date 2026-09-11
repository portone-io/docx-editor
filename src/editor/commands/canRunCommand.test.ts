// @vitest-environment jsdom
import { deleteSelection } from "prosemirror-commands";
import type { Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx, makeNotesDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import {
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  mergeCells,
} from "../../table/commands";
import { createEditorState } from "../createEditor";
import { canRunCommand } from "./canRunCommand";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const LOCKED_PR =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

const cell = (inner: string) => `<w:tc><w:p>${inner}</w:p></w:tc>`;

const lockedCell = (text: string) =>
  cell(`<w:sdt>${LOCKED_PR}<w:sdtContent>${run(text)}</w:sdtContent></w:sdt>`);

/** A two by two table whose top left cell holds a locked control */
const BODY =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${lockedCell("Locked")}${cell(run("TopRight"))}</w:tr>` +
  `<w:tr>${cell(run("BottomLeft"))}${cell(run("BottomRight"))}</w:tr>` +
  "</w:tbl>";

function posOfText(doc: PMNode, needle: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found < 0 && node.isText && node.text === needle) found = pos;
  });
  if (found < 0) throw new Error(`text not found: ${needle}`);
  return found;
}

function caretAt(needle: string): EditorState {
  const state = createEditorState(importDocx(makeDocx(BODY)).doc);
  const at = posOfText(state.doc, needle);
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, at))
  );
}

describe("a table command that would reach into locked content", () => {
  /**
   * The command answers for the guard itself (`table/commands`), so this says the same thing twice.
   * It is asserted all the same, because the two answering differently is the very fault this file
   * exists for.
   */
  it("is reported as one that cannot be run, and reports so itself", () => {
    const state = caretAt("Locked");
    expect(deleteRow(state)).toBe(false);
    expect(canRunCommand(deleteRow, state)).toBe(false);
    expect(canRunCommand(deleteColumn, state)).toBe(false);
    expect(canRunCommand(deleteTable, state)).toBe(false);
  });

  /** The row is what is deleted, so a lock anywhere in it counts */
  it("is reported the same from the cell beside the locked one", () => {
    expect(canRunCommand(deleteRow, caretAt("TopRight"))).toBe(false);
    expect(canRunCommand(deleteColumn, caretAt("TopRight"))).toBe(true);
  });

  it("leaves the rows and columns with no lock in them as they were", () => {
    expect(canRunCommand(deleteRow, caretAt("BottomLeft"))).toBe(true);
    expect(canRunCommand(deleteColumn, caretAt("BottomRight"))).toBe(true);
    expect(canRunCommand(deleteTable, caretAt("BottomRight"))).toBe(false);
  });

  it("does not stand in the way of inserting a row beside a lock", () => {
    expect(canRunCommand(addRowBefore, caretAt("Locked"))).toBe(true);
  });

  it("is reported as unavailable where the command itself does not apply", () => {
    expect(canRunCommand(mergeCells, caretAt("Locked"))).toBe(false);
  });
});

/** One paragraph reading "mkn" with a bookmark range anchored around the "k" */
const BOOKMARK_P =
  "<w:p>" +
  run("m") +
  '<w:bookmarkStart w:id="9" w:name="b"/>' +
  run("k") +
  '<w:bookmarkEnd w:id="9"/>' +
  run("n") +
  "</w:p>";

/** A selection running from the character before the first node of this type to the one after it */
function across(state: EditorState, typeName: string): EditorState {
  const spans: { pos: number; size: number }[] = [];
  state.doc.descendants((node, pos) => {
    if (spans.length === 0 && node.type.name === typeName) {
      spans.push({ pos, size: node.nodeSize });
    }
    return spans.length === 0;
  });
  const span = spans[0];
  if (span === undefined) throw new Error(`no ${typeName} in the document`);
  return state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, span.pos - 1, span.pos + span.size + 1)
    )
  );
}

/** The state a command answers about, and the state dispatching it really leaves behind */
function ran(state: EditorState, command: Command): EditorState {
  let after = state;
  command(state, (tr) => {
    after = after.apply(tr);
  });
  return after;
}

/**
 * A bookmark marker and an endnote reference are preserved rather than edited, so a command that
 * would sweep one away is refused when it is dispatched. The answer has to say so beforehand.
 */
describe("a command that would sweep away a preserved marker", () => {
  it("reports false for deleteSelection across an endnote reference", () => {
    const opened = createEditorState(
      importDocx(
        makeNotesDocx(
          `<w:p>${run("Text")}<w:r><w:endnoteReference w:id="3"/></w:r>${run(" more")}</w:p>`
        )
      ).doc
    );
    const state = across(opened, "noteReference");

    expect(canRunCommand(deleteSelection, state)).toBe(false);
    expect(ran(state, deleteSelection).doc.eq(state.doc)).toBe(true);
  });

  it("reports true for deleteSelection across a footnote reference, which may go", () => {
    const opened = createEditorState(
      importDocx(
        makeNotesDocx(
          `<w:p>${run("Text")}<w:r><w:footnoteReference w:id="2"/></w:r>${run(" more")}</w:p>`
        )
      ).doc
    );
    const state = across(opened, "noteReference");

    expect(canRunCommand(deleteSelection, state)).toBe(true);
    expect(ran(state, deleteSelection).doc.textContent).toBe("Texmore");
  });

  it("reports false for deleteSelection across a bookmark marker", () => {
    const opened = createEditorState(importDocx(makeDocx(BOOKMARK_P)).doc);
    const state = across(opened, "rawInline");

    expect(canRunCommand(deleteSelection, state)).toBe(false);
    expect(ran(state, deleteSelection).doc.eq(state.doc)).toBe(true);
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
function startComposition(live: EditorView): void {
  live.dom.dispatchEvent(
    new CompositionEvent("compositionstart", { bubbles: true, data: "" })
  );
}

/** A refusal is answered on the frame after it, so two frames are waited out */
function nextFrame(): Promise<void> {
  return new Promise((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
  });
}

describe("asking whether a command can be run", () => {
  it("leaves an open composition standing where the guard says no", async () => {
    const live = mounted(caretAt("Locked"));
    startComposition(live);
    expect(live.composing).toBe(true);

    const before = live.state;
    expect(canRunCommand(deleteRow, live.state)).toBe(false);

    await nextFrame();
    expect(live.composing).toBe(true);
    expect(live.state).toBe(before);
  });

  /** The second transaction stands on the document the first produced, not on the state's own */
  it("answers a command that dispatches twice", () => {
    const insertsTwice: Command = (state, dispatch) => {
      if (!dispatch) return true;
      const first = state.tr.insertText("one", state.selection.from);
      dispatch(first);
      const after = state.apply(first);
      dispatch(after.tr.insertText("two", after.selection.from));
      return true;
    };
    expect(canRunCommand(insertsTwice, caretAt("BottomRight"))).toBe(true);
  });
});
