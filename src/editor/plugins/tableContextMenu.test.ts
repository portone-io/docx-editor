// @vitest-environment jsdom
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import type { EditingProtection } from "../../schema/protection";
import { createEditorState, createEditorView } from "../createEditor";
import { closeTableMenu, tableMenuAnchor } from "./tableContextMenu";

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

function cellWithText(view: { dom: Element }, text: string): Element {
  const cell = Array.from(view.dom.querySelectorAll("td")).find(
    (td) => td.textContent === text
  );
  if (!cell) throw new Error(`cell not found: ${text}`);
  return cell;
}

/**
 * Right clicks the target, with the spot the click lands on handed in.
 * The text menu stands in front of this plugin and works that spot out from the coordinates,
 * which jsdom draws nothing for, so the answer is handed over here instead.
 */
function rightClick(
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

/** The position of the first character of the text this cell holds */
function textStartIn(view: EditorView, cellText: string): number {
  return view.posAtDOM(cellWithText(view, cellText), 0) + 1;
}

describe("right clicking a table cell", () => {
  it("blocks the browser menu inside a cell and reports where the menu goes", () => {
    const view = openEditor();
    const cell = cellWithText(view, "Left");
    expect(rightClick(view, cell, textStartIn(view, "Left"))).toBe(true);
    expect(tableMenuAnchor(view.state)).toEqual({ clientX: 120, clientY: 240 });
  });

  /** The text menu is what answers a click in the body, so this menu has no anchor there */
  it("opens no table menu in the body text", () => {
    const view = openEditor();
    const paragraph = view.dom.querySelector("p");
    if (!paragraph) throw new Error("paragraph not found");
    rightClick(view, paragraph, 2);
    expect(tableMenuAnchor(view.state)).toBeNull();
  });

  it("does not block when readOnly", () => {
    const view = openEditor("readOnly");
    const cell = cellWithText(view, "Left");
    expect(rightClick(view, cell, textStartIn(view, "Left"))).toBe(false);
    expect(tableMenuAnchor(view.state)).toBeNull();
  });

  /** The text menu takes the click instead, which `textContextMenu.test.ts` covers */
  it("stays shut for a commenter, who may do nothing to a table", () => {
    const view = openEditor("comments");
    rightClick(view, cellWithText(view, "Left"), textStartIn(view, "Left"));
    expect(tableMenuAnchor(view.state)).toBeNull();
  });

  it("moves the selection into the cell when clicking a cell that did not hold the cursor", () => {
    const view = openEditor();
    rightClick(view, cellWithText(view, "Right"), textStartIn(view, "Right"));
    expect(view.state.selection.$from.parent.textContent).toBe("Right");
  });

  it("leaves the selection alone when the cursor is already in that cell", () => {
    const view = openEditor();
    const landing = textStartIn(view, "Right");
    rightClick(view, cellWithText(view, "Right"), landing);
    const selected = view.state.selection;
    rightClick(view, cellWithText(view, "Right"), landing);
    expect(view.state.selection.from).toBe(selected.from);
  });

  it("closing clears the anchor, and there is nothing left to do once it is closed", () => {
    const view = openEditor();
    rightClick(view, cellWithText(view, "Left"), textStartIn(view, "Left"));
    expect(closeTableMenu(view.state, (tr) => view.dispatch(tr))).toBe(true);
    expect(tableMenuAnchor(view.state)).toBeNull();
    expect(closeTableMenu(view.state)).toBe(false);
  });

  it("closes by itself when the document changes", () => {
    const view = openEditor();
    rightClick(view, cellWithText(view, "Left"), textStartIn(view, "Left"));
    view.dispatch(view.state.tr.insertText("x"));
    expect(tableMenuAnchor(view.state)).toBeNull();
  });
});
