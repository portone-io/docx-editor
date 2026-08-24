// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
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
