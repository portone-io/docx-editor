// @vitest-environment jsdom
import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { AUTHOR, EDITING } from "../__testing__/mode";
import { renderInto } from "../__testing__/react";
import {
  DocxEditor,
  type DocxEditorHandle,
  type DocxEditorMode,
} from "../DocxEditor";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

/**
 * jsdom draws nothing, and the editor asks where the caret is every time it scrolls a change
 * into view. Empty rectangles are answer enough, and they are borrowed off an element so that
 * they are the very shapes the browser would hand back.
 */
const noRects = document.createElement("div").getClientRects();
const emptyRect = document.createElement("div").getBoundingClientRect();
Range.prototype.getClientRects = () => noRects;
Range.prototype.getBoundingClientRect = () => emptyRect;

let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  host.remove();
});

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const LOCKED_PR =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

const cell = (inner: string) => `<w:tc><w:p>${inner}</w:p></w:tc>`;

/** A two by two table whose top left cell holds a locked control */
const WITH_LOCKED_CELL = makeDocx(
  "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
    "<w:tr>" +
    cell(
      `<w:sdt>${LOCKED_PR}<w:sdtContent>${run("Locked")}</w:sdtContent></w:sdt>`
    ) +
    cell(run("TopRight")) +
    "</w:tr>" +
    `<w:tr>${cell(run("BottomLeft"))}${cell(run("BottomRight"))}</w:tr>` +
    "</w:tbl>"
);

/** The same table with the control wrapped around the whole top left cell instead */
const WITH_LOCKED_WRAPPER = makeDocx(
  "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
    "<w:tr>" +
    `<w:sdt>${LOCKED_PR}<w:sdtContent>${cell(run("Locked"))}</w:sdtContent></w:sdt>` +
    cell(run("TopRight")) +
    "</w:tr>" +
    `<w:tr>${cell(run("BottomLeft"))}${cell(run("BottomRight"))}</w:tr>` +
    "</w:tbl>"
);

/** The same again, with nothing written inside the cell the lock wraps */
const WITH_EMPTY_LOCKED_WRAPPER = makeDocx(
  "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
    "<w:tr>" +
    `<w:sdt>${LOCKED_PR}<w:sdtContent>${cell("")}</w:sdtContent></w:sdt>` +
    cell(run("TopRight")) +
    "</w:tr>" +
    `<w:tr>${cell(run("BottomLeft"))}${cell(run("BottomRight"))}</w:tr>` +
    "</w:tbl>"
);

const WITHOUT_LOCKED_CELL = makeDocx(
  `<w:p>${run("Before")}</w:p>` +
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
    `<w:tr>${cell(run("Free"))}</w:tr>` +
    "</w:tbl>"
);

/** The mode a screen where a template is authored mounts the editor in */
const AUTHORING: DocxEditorMode = {
  kind: "edit",
  author: AUTHOR,
  locking: true,
};

const render = (element: ReactNode) => renderInto(host, element);

function mount(
  bytes: Uint8Array,
  { mode = EDITING }: { mode?: DocxEditorMode } = {}
) {
  const box: { current: DocxEditorHandle | null } = { current: null };
  const unmount = render(
    <DocxEditor
      document={bytes}
      ref={box}
      renderImportError={() => null}
      mode={mode}
    />
  );
  const handle = box.current;
  if (!handle) throw new Error("the ref was not attached");
  // jsdom draws nothing, so the spot a right click lands on is answered here: the caret the
  // state already holds, which is what a click on the current selection lands on
  handle.view.posAtCoords = () => ({
    pos: handle.view.state.selection.head,
    inside: -1,
  });
  return { handle, unmount };
}

function rightClickCell(text: string): void {
  const found = Array.from(host.querySelectorAll("td")).find(
    (td) => td.textContent === text
  );
  if (!found) throw new Error(`cell not found: ${text}`);
  act(() => {
    found.dispatchEvent(
      new MouseEvent("contextmenu", {
        bubbles: true,
        cancelable: true,
        clientX: 40,
        clientY: 60,
      })
    );
  });
}

/** The text written on every row of the open menu */
function labels(): string[] {
  return Array.from(host.querySelectorAll('button[role="menuitem"]')).map(
    (element) => element.textContent ?? ""
  );
}

/** A row of the open menu, found by the text written on it */
function item(label: string): HTMLButtonElement {
  const found = Array.from(host.querySelectorAll('button[role="menuitem"]'))
    .filter(
      (element): element is HTMLButtonElement =>
        element instanceof HTMLButtonElement
    )
    .find((element) => element.textContent === label);
  if (!found) throw new Error(`menu item not found: ${label}`);
  return found;
}

/** Whether a row is drawn with nothing to do. The rows stay focusable, so the state is an aria one */
function blocked(label: string): boolean {
  return item(label).getAttribute("aria-disabled") === "true";
}

/** The row the focus is on, by the text written on it */
function focusedRow(): string | null {
  const found = document.activeElement;
  return found?.getAttribute("role") === "menuitem"
    ? (found.textContent ?? "")
    : null;
}

/** The one row of the menu that Tab can reach. The rest are walked with the arrows */
function tabbableRows(): string[] {
  return Array.from(host.querySelectorAll('button[role="menuitem"]'))
    .filter((element) => element.getAttribute("tabindex") === "0")
    .map((element) => element.textContent ?? "");
}

function press(key: string): void {
  const target = document.activeElement ?? document.body;
  act(() => {
    target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
  });
}

describe("the merge cells row", () => {
  it("is not clickable before several cells are selected", () => {
    const { unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");

    expect(blocked("Merge cells")).toBe(true);
    unmount();
  });
});

/**
 * A row with nothing to do stays in the tab order rather than being disabled outright, so the
 * refusal has to be the click's own.
 */
describe("a row drawn with nothing to do", () => {
  it("stays focusable and does nothing at all when it is clicked", () => {
    const { handle, unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("Locked");
    const before = handle.view.state.doc.child(0).childCount;
    const row = item("Delete row");
    expect(blocked("Delete row")).toBe(true);

    row.focus();
    expect(document.activeElement).toBe(row);

    act(() => row.click());
    expect(handle.view.state.doc.child(0).childCount).toBe(before);
    // The menu is still standing, since nothing was chosen
    expect(labels()).toContain("Delete row");
    unmount();
  });
});

describe("the delete table row", () => {
  it("is unavailable when deleting the table would remove locked content", () => {
    const { unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");

    expect(blocked("Delete table")).toBe(true);
    unmount();
  });

  it("removes an unlocked table and closes the menu", () => {
    const { handle, unmount } = mount(WITHOUT_LOCKED_CELL);
    rightClickCell("Free");

    expect(blocked("Delete table")).toBe(false);
    act(() => item("Delete table").click());

    expect(handle.view.state.doc.textContent).toBe("Before");
    expect(host.querySelector("table")).toBeNull();
    expect(labels()).toHaveLength(0);
    unmount();
  });
});

describe("the table menu over a cell holding locked text", () => {
  it("keeps the rows that would rewrite the lock unclickable", () => {
    const { unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("Locked");

    expect(blocked("Delete row")).toBe(true);
    expect(blocked("Delete column")).toBe(true);
    // Making room around the lock changes nothing inside it
    expect(blocked("Insert row above")).toBe(false);
    expect(blocked("Insert column left")).toBe(false);
    unmount();
  });

  it("keeps them unclickable over a cell the lock wraps whole", () => {
    const { unmount } = mount(WITH_LOCKED_WRAPPER);
    rightClickCell("Locked");

    expect(blocked("Delete row")).toBe(true);
    expect(blocked("Delete column")).toBe(true);
    expect(blocked("Insert row above")).toBe(false);
    unmount();
  });

  it("leaves the rows over a cell with no lock in reach clickable", () => {
    const { handle, unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");

    expect(blocked("Delete row")).toBe(false);
    const rowsBefore = handle.view.state.doc.child(0).childCount;
    act(() => item("Delete row").click());
    expect(handle.view.state.doc.child(0).childCount).toBe(rowsBefore - 1);
    unmount();
  });
});

/**
 * A cell shut as a whole is never reached by the text menu, which leaves this menu as the one way
 * to lift such a lock
 */
describe("the unlocking row of the table menu", () => {
  it("is not drawn unless the consumer turned locking on", () => {
    const { unmount } = mount(WITH_LOCKED_WRAPPER);
    rightClickCell("Locked");

    expect(labels()).not.toContain("Unlock");
    unmount();
  });

  /** Every other row stands whatever the cell holds, so this one does too */
  it("stands unclickable over a cell with no lock to lift", () => {
    const { unmount } = mount(WITH_LOCKED_WRAPPER, { mode: AUTHORING });
    rightClickCell("BottomRight");

    expect(labels()).toContain("Unlock");
    expect(blocked("Unlock")).toBe(true);
    unmount();
  });

  it("lifts the lock off the clicked cell and opens it to editing", () => {
    const { handle, unmount } = mount(WITH_LOCKED_WRAPPER, { mode: AUTHORING });
    rightClickCell("Locked");
    expect(blocked("Unlock")).toBe(false);

    act(() => item("Unlock").click());
    rightClickCell("Locked");
    expect(blocked("Unlock")).toBe(true);
    // The row that the lock had shut is clickable again
    expect(blocked("Delete row")).toBe(false);
    // The control Word wrote around the cell is still standing
    const cellNode = handle.view.state.doc.child(0).child(0).child(0);
    expect(cellNode.attrs.sdtContentsLocked).toBe(false);
    expect(cellNode.attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>'
    );
    unmount();
  });

  /**
   * A cell with nothing in it has no text to select either, so this menu is the whole of the way
   * to a lock wrapped around it
   */
  it("lifts the lock off a cell that holds no text at all", () => {
    const { handle, unmount } = mount(WITH_EMPTY_LOCKED_WRAPPER, {
      mode: AUTHORING,
    });
    rightClickCell("");
    expect(blocked("Unlock")).toBe(false);

    act(() => item("Unlock").click());
    const cellNode = handle.view.state.doc.child(0).child(0).child(0);
    expect(cellNode.attrs.sdtContentsLocked).toBe(false);

    rightClickCell("");
    expect(blocked("Unlock")).toBe(true);
    // What the lock had shut is open again
    expect(blocked("Delete row")).toBe(false);
    unmount();
  });
});

/**
 * Inside a table Tab belongs to the table, so a menu that left the focus on the paper could only
 * ever be reached with a mouse. It takes the focus as it opens instead, and answers the keys of
 * the menu pattern from there.
 */
describe("the table menu on a keyboard", () => {
  it("takes the focus onto its first row as it opens", () => {
    const { unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");

    expect(focusedRow()).toBe("Insert row above");
    // The rows are one stop in the page's tab order rather than one each
    expect(tabbableRows()).toEqual(["Insert row above"]);
    unmount();
  });

  it("walks the rows with the arrows, wrapping at both ends", () => {
    const { unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");

    press("ArrowDown");
    expect(focusedRow()).toBe("Insert row below");
    press("ArrowUp");
    expect(focusedRow()).toBe("Insert row above");
    press("ArrowUp");
    expect(focusedRow()).toBe("Delete table");
    unmount();
  });

  it("jumps to the ends with Home and End", () => {
    const { unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");

    press("End");
    expect(focusedRow()).toBe("Delete table");
    press("Home");
    expect(focusedRow()).toBe("Insert row above");
    unmount();
  });

  it("jumps to the next row starting with the letter typed", () => {
    const { unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");

    press("d");
    expect(focusedRow()).toBe("Delete row");
    press("d");
    expect(focusedRow()).toBe("Delete column");
    press("d");
    expect(focusedRow()).toBe("Delete table");
    unmount();
  });

  it("walks a row with nothing to do like any other, so it is heard rather than skipped", () => {
    const { unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("Locked");

    press("ArrowDown");
    press("ArrowDown");
    expect(focusedRow()).toBe("Delete row");
    expect(blocked("Delete row")).toBe(true);
    unmount();
  });

  it("runs the row the keyboard walked to", () => {
    const { handle, unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");
    const rowsBefore = handle.view.state.doc.child(0).childCount;

    press("ArrowDown");
    press("ArrowDown");
    expect(focusedRow()).toBe("Delete row");
    act(() => item("Delete row").click());

    expect(handle.view.state.doc.child(0).childCount).toBe(rowsBefore - 1);
    expect(labels()).toHaveLength(0);
    unmount();
  });

  it("closes on Tab and hands the focus back to the paper", () => {
    const { handle, unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");
    expect(focusedRow()).toBe("Insert row above");

    press("Tab");
    expect(labels()).toHaveLength(0);
    expect(document.activeElement).toBe(handle.view.dom);
    unmount();
  });

  it("closes on Escape and hands the focus back to the paper", () => {
    const { handle, unmount } = mount(WITH_LOCKED_CELL);
    rightClickCell("BottomRight");

    press("Escape");
    expect(labels()).toHaveLength(0);
    expect(document.activeElement).toBe(handle.view.dom);
    unmount();
  });
});
