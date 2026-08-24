// @vitest-environment jsdom
import { unzipSync } from "fflate";
import { TextSelection } from "prosemirror-state";
import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { decode, makeDocx } from "../__testing__/docx";
import { renderInto } from "../__testing__/react";
import {
  DocxEditor,
  type DocxEditorHandle,
  type DocxEditorMode,
} from "../DocxEditor";
import { toRunFormat } from "../model/format";

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
/** The clipboard commands handed to the browser, which jsdom has none of */
let handedOver: string[];

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
  handedOver = [];
  document.execCommand = (command: string) => {
    handedOver.push(command);
    return true;
  };
});

afterEach(() => {
  host.remove();
});

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const PARAGRAPH = makeDocx(`<w:p>${run("Source")}</w:p>`);

const LOCKED_PR =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

/** Plain text, a locked control, then plain text again */
const WITH_LOCK = makeDocx(
  `<w:p>${run("pre")}<w:sdt>${LOCKED_PR}<w:sdtContent>${run("locked")}` +
    `</w:sdtContent></w:sdt>${run("post")}</w:p>`
);

const cellXml = (text: string) => `<w:tc><w:p>${run(text)}</w:p></w:tc>`;

const WITH_TABLE = makeDocx(
  `<w:p>${run("Body")}</w:p>` +
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
    `<w:tr>${cellXml("Left")}</w:tr>` +
    "</w:tbl>"
);

/** The mode a screen where a template is authored mounts the editor in */
const AUTHORING: DocxEditorMode = {
  kind: "edit",
  locking: true,
};

const render = (element: ReactNode) => renderInto(host, element);

function mount(bytes: Uint8Array, props: { mode?: DocxEditorMode } = {}) {
  const box: { current: DocxEditorHandle | null } = { current: null };
  const unmount = render(
    <DocxEditor
      document={bytes}
      ref={box}
      renderImportError={() => null}
      {...props}
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

/** Right clicks the first paragraph on screen */
function rightClickText(): void {
  const paragraph = host.querySelector("p");
  if (!paragraph) throw new Error("paragraph not found");
  act(() => {
    paragraph.dispatchEvent(
      new MouseEvent("contextmenu", {
        bubbles: true,
        cancelable: true,
        clientX: 40,
        clientY: 60,
      })
    );
  });
}

/** A row of the open menu. The shortcut reminder follows the label, so the label is a prefix */
function item(label: string): HTMLButtonElement {
  const found = Array.from(host.querySelectorAll('button[role="menuitem"]'))
    .filter(
      (element): element is HTMLButtonElement =>
        element instanceof HTMLButtonElement
    )
    .find((element) => element.textContent?.startsWith(label));
  if (!found) throw new Error(`menu item not found: ${label}`);
  return found;
}

/** Whether a row is drawn with nothing to do. The rows stay focusable, so the state is an aria one */
function blocked(label: string): boolean {
  return item(label).getAttribute("aria-disabled") === "true";
}

function labels(): string[] {
  return Array.from(host.querySelectorAll('button[role="menuitem"]')).map(
    (element) => element.textContent ?? ""
  );
}

function select(handle: DocxEditorHandle, from: number, to: number): void {
  act(() => {
    const { view } = handle;
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to))
    );
  });
}

function documentXml(handle: DocxEditorHandle): string {
  return decode(unzipSync(handle.exportBytes())["word/document.xml"]);
}

describe("the text right click menu", () => {
  it("offers the clipboard entries with their shortcuts", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    select(handle, 1, 3);
    rightClickText();

    // jsdom is no Mac, so the reminders are spelled out
    expect(labels()).toEqual([
      "CutCtrl+X",
      "CopyCtrl+C",
      "PasteCtrl+V",
      "Delete",
      "Add comment",
    ]);
    unmount();
  });

  it("keeps cut, copy and delete unclickable while nothing is selected", () => {
    const { unmount } = mount(PARAGRAPH);
    rightClickText();

    expect(blocked("Cut")).toBe(true);
    expect(blocked("Copy")).toBe(true);
    expect(blocked("Delete")).toBe(true);
    expect(blocked("Add comment")).toBe(true);
    // There is a caret to paste at even with nothing selected
    expect(blocked("Paste")).toBe(false);
    unmount();
  });

  it.each(["Cut", "Copy"])("hands %s to the browser and closes", (label) => {
    const { handle, unmount } = mount(PARAGRAPH);
    select(handle, 1, 3);
    rightClickText();
    act(() => item(label).click());

    expect(handedOver).toEqual([label.toLowerCase()]);
    expect(host.querySelector('[role="menu"]')).toBeNull();
    unmount();
  });

  it("deletes the selected text", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    select(handle, 1, 7);
    rightClickText();
    act(() => item("Delete").click());

    expect(handle.view.state.doc.child(0).textContent).toBe("");
    unmount();
  });

  it("pastes the clipboard text with its formatting dropped", async () => {
    Object.defineProperty(navigator, "clipboard", {
      value: { readText: () => Promise.resolve("Pasted") },
      configurable: true,
    });
    const { handle, unmount } = mount(PARAGRAPH);
    select(handle, 1, 1);
    rightClickText();
    await act(async () => {
      item("Paste").click();
    });

    expect(handle.view.state.doc.child(0).textContent).toBe("PastedSource");
    unmount();
  });

  it("uses rich clipboard data when the browser makes it available", async () => {
    Object.defineProperty(navigator, "clipboard", {
      value: {
        read: () =>
          Promise.resolve([
            {
              types: ["text/html", "text/plain"],
              getType: (type: string) =>
                Promise.resolve({
                  text: () =>
                    Promise.resolve(
                      type === "text/html"
                        ? '<strong style="font-family: Arial; color: #123456">Pasted</strong>'
                        : "Pasted"
                    ),
                }),
            },
          ]),
        readText: () => Promise.resolve("Pasted"),
      },
      configurable: true,
    });
    const { handle, unmount } = mount(PARAGRAPH);
    select(handle, 1, 1);
    rightClickText();
    await act(async () => {
      item("Paste").click();
    });

    const pasted = handle.view.state.doc.child(0).child(0);
    const run = pasted.marks.find((mark) => mark.type.name === "run");
    expect(pasted.textContent).toBe("Pasted");
    expect(toRunFormat(run?.attrs.format)).toMatchObject({
      bold: true,
      fontFamily: '"Arial"',
      color: "#123456",
    });
    unmount();
  });

  it("does not show up on a cell with nothing selected, where the table menu belongs", () => {
    const { unmount } = mount(WITH_TABLE);
    const cell = host.querySelector("td");
    if (!cell) throw new Error("cell not found");
    act(() => {
      cell.dispatchEvent(
        new MouseEvent("contextmenu", { bubbles: true, cancelable: true })
      );
    });

    expect(labels()).toContain("Insert row below");
    // Not one row of the text menu is drawn beside it
    for (const row of ["CutCtrl+X", "CopyCtrl+C", "PasteCtrl+V", "Delete"]) {
      expect(labels()).not.toContain(row);
    }
    unmount();
  });
});

describe("the lock entries", () => {
  it("are not drawn unless the consumer turned locking on", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    select(handle, 1, 3);
    rightClickText();

    expect(labels()).not.toContain("Lock");
    expect(host.querySelectorAll("hr")).toHaveLength(1);
    unmount();
  });

  it("are set off by a separator once locking is on", () => {
    const { handle, unmount } = mount(PARAGRAPH, { mode: AUTHORING });
    select(handle, 1, 3);
    rightClickText();

    expect(labels()).toEqual([
      "CutCtrl+X",
      "CopyCtrl+C",
      "PasteCtrl+V",
      "Delete",
      "Add comment",
      "Lock",
    ]);
    expect(host.querySelectorAll("hr")).toHaveLength(2);
    unmount();
  });

  it("locking writes a shut control into the document", () => {
    const { handle, unmount } = mount(PARAGRAPH, { mode: AUTHORING });
    select(handle, 1, 3);
    rightClickText();
    // Away from any lock the one row offers locking
    expect(labels()).not.toContain("Unlock");
    expect(blocked("Lock")).toBe(false);

    act(() => item("Lock").click());
    expect(documentXml(handle)).toContain('<w:lock w:val="sdtContentLocked"/>');
    unmount();
  });

  it("what is locked can be neither cut nor pasted over, and unlocking opens it again", () => {
    const { handle, unmount } = mount(WITH_LOCK, { mode: AUTHORING });
    // A caret in the middle of the locked text
    select(handle, 7, 7);
    rightClickText();

    expect(blocked("Paste")).toBe(true);
    // Over the locked stretch the same row turns into unlocking
    expect(labels()).not.toContain("Lock");
    expect(blocked("Unlock")).toBe(false);

    act(() => item("Unlock").click());
    // The control itself stays standing; only its lock goes
    expect(documentXml(handle)).toContain("<w:sdt>");
    expect(documentXml(handle)).not.toContain("<w:lock");

    select(handle, 4, 10);
    rightClickText();
    expect(blocked("Cut")).toBe(false);
    unmount();
  });

  // What is already settled outranks settling more, so the row offers the lifting of the two
  it("offer lifting where the selection holds both a lock and text to lock", () => {
    const { handle, unmount } = mount(WITH_LOCK, { mode: AUTHORING });
    select(handle, 1, 14);
    rightClickText();

    expect(labels()).not.toContain("Lock");
    expect(blocked("Unlock")).toBe(false);
    unmount();
  });

  it("cutting is unclickable where the selection runs over locked text", () => {
    const { handle, unmount } = mount(WITH_LOCK, { mode: AUTHORING });
    select(handle, 1, 14);
    rightClickText();

    expect(blocked("Cut")).toBe(true);
    expect(blocked("Delete")).toBe(true);
    expect(blocked("Copy")).toBe(false);
    unmount();
  });
});
