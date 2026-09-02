// @vitest-environment jsdom
import { TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import type { EditingProtection } from "../../schema/protection";
import { createEditorState, createEditorView } from "../createEditor";
import { tableMenuAnchor } from "./tableContextMenu";
import { closeTextMenu, textMenuAnchor } from "./textContextMenu";

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

const BODY =
  '<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>' +
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cellXml("Left")}${cellXml("Right")}</w:tr>` +
  "</w:tbl>";

let mounted: (() => void)[] = [];

afterEach(() => {
  for (const dispose of mounted) dispose();
  mounted = [];
});

function openEditor(protection: EditingProtection = "none"): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  const { doc, session } = importDocx(makeDocx(BODY));
  const view = createEditorView({
    mount,
    state: createEditorState(doc, { protection }),
    defaults: session.defaults,
    onStateChange: () => undefined,
  });
  mounted.push(() => {
    view.destroy();
    mount.remove();
  });
  return view;
}

/**
 * Right clicks the target, with the spot the click lands on handed in.
 * jsdom draws nothing, so the coordinate lookup the plugin makes is answered here instead, the
 * way a browser would answer it.
 */
function rightClickAt(
  view: EditorView,
  target: Element,
  landing: number
): boolean {
  view.posAtCoords = () => ({ pos: landing, inside: -1 });
  const event = new MouseEvent("contextmenu", {
    bubbles: true,
    cancelable: true,
    clientX: 120,
    clientY: 240,
  });
  target.dispatchEvent(event);
  return event.defaultPrevented;
}

function selectText(view: EditorView, from: number, to: number): void {
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to))
  );
}

function bodyParagraph(view: EditorView): Element {
  const paragraph = view.dom.querySelector("p");
  if (!paragraph) throw new Error("paragraph not found");
  return paragraph;
}

function cellWithText(view: EditorView, text: string): Element {
  const cell = Array.from(view.dom.querySelectorAll("td")).find(
    (td) => td.textContent === text
  );
  if (!cell) throw new Error(`cell not found: ${text}`);
  return cell;
}

/** The position of the first character of the text this cell holds */
function textStartIn(view: EditorView, cellText: string): number {
  return view.posAtDOM(cellWithText(view, cellText), 0) + 1;
}

describe("right clicking the text", () => {
  it("blocks the browser menu and reports where the menu goes", () => {
    const view = openEditor();
    expect(rightClickAt(view, bodyParagraph(view), 2)).toBe(true);
    expect(textMenuAnchor(view.state)).toEqual({ clientX: 120, clientY: 240 });
  });

  it("does not block when readOnly", () => {
    const view = openEditor("readOnly");
    expect(rightClickAt(view, bodyParagraph(view), 2)).toBe(false);
    expect(textMenuAnchor(view.state)).toBeNull();
  });

  /** With nothing selected, a cell is right clicked to reach the row and column actions */
  it("leaves a cell with nothing selected to the table menu", () => {
    const view = openEditor();
    rightClickAt(view, cellWithText(view, "Left"), textStartIn(view, "Left"));
    expect(textMenuAnchor(view.state)).toBeNull();
    expect(tableMenuAnchor(view.state)).not.toBeNull();
  });

  /** A block of selected cells is how merging is reached, so it stays the table menu's click */
  it("leaves a cell with whole cells selected to the table menu", () => {
    const view = openEditor();
    const first = view.posAtDOM(cellWithText(view, "Left"), 0);
    const second = view.posAtDOM(cellWithText(view, "Right"), 0);
    view.dispatch(
      view.state.tr.setSelection(
        CellSelection.create(view.state.doc, first - 1, second - 1)
      )
    );

    rightClickAt(view, cellWithText(view, "Left"), textStartIn(view, "Left"));
    expect(textMenuAnchor(view.state)).toBeNull();
    expect(tableMenuAnchor(view.state)).not.toBeNull();
  });

  it("takes a cell where the click lands on selected text, so the clipboard is at hand", () => {
    const view = openEditor();
    const cell = cellWithText(view, "Left");
    const at = textStartIn(view, "Left");
    selectText(view, at, at + 2);

    expect(rightClickAt(view, cell, at + 1)).toBe(true);
    expect(textMenuAnchor(view.state)).not.toBeNull();
    expect(tableMenuAnchor(view.state)).toBeNull();
  });

  it("leaves a cell to the table menu while the selected text stands elsewhere", () => {
    const view = openEditor();
    selectText(view, 1, 3);

    rightClickAt(view, cellWithText(view, "Left"), textStartIn(view, "Left"));
    expect(textMenuAnchor(view.state)).toBeNull();
    expect(tableMenuAnchor(view.state)).not.toBeNull();
    expect(view.state.selection.$from.parent.textContent).toBe("Left");
  });

  it("takes the caret along to the spot the click landed on", () => {
    const view = openEditor();
    selectText(view, 1, 1);

    rightClickAt(view, bodyParagraph(view), 3);
    expect(view.state.selection.empty).toBe(true);
    expect(view.state.selection.from).toBe(3);
  });

  it("leaves the selected text alone when the click lands inside it", () => {
    const view = openEditor();
    selectText(view, 1, 3);

    rightClickAt(view, bodyParagraph(view), 2);
    expect(view.state.selection.from).toBe(1);
    expect(view.state.selection.to).toBe(3);
  });

  it("closing clears the anchor, and there is nothing left to do once it is closed", () => {
    const view = openEditor();
    rightClickAt(view, bodyParagraph(view), 2);
    expect(closeTextMenu(view.state, (tr) => view.dispatch(tr))).toBe(true);
    expect(textMenuAnchor(view.state)).toBeNull();
    expect(closeTextMenu(view.state)).toBe(false);
  });

  it("closes by itself when the document changes", () => {
    const view = openEditor();
    rightClickAt(view, bodyParagraph(view), 2);
    view.dispatch(view.state.tr.insertText("x"));
    expect(textMenuAnchor(view.state)).toBeNull();
  });
});
