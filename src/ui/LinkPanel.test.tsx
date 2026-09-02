// @vitest-environment jsdom
/**
 * The link panel driven the way a person drives it: the key both editors bind, the toolbar button,
 * and the keyboard once the panel stands open.
 *
 * What the commands behind it do is `editor/commands/linkCommands.test.ts`; what is checked here is
 * that the panel opens where it should, shows the address that is there, hands the focus back on
 * closing, and never offers what the commands would refuse.
 */

import { unzipSync } from "fflate";
import { TextSelection } from "prosemirror-state";
import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { decode, makeDocx, makeLinkedDocx } from "../__testing__/docx";
import { EDITING } from "../__testing__/mode";
import { renderInto } from "../__testing__/react";
import { DocxEditor, type DocxEditorHandle } from "../DocxEditor";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

const TERMS = "https://example.com/terms";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const PLAIN = makeDocx(`<w:p>${run("see our terms")}</w:p>`);

const LINKED = makeLinkedDocx(
  `<w:p>${run("see ")}<w:hyperlink r:id="rId9">${run("our terms")}` +
    "</w:hyperlink></w:p>",
  { rId9: TERMS }
);

let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  host.remove();
});

const render = (element: ReactNode) => renderInto(host, element);

function mount(bytes: Uint8Array) {
  const box: { current: DocxEditorHandle | null } = { current: null };
  const unmount = render(
    <DocxEditor
      document={bytes}
      mode={EDITING}
      ref={box}
      renderImportError={() => null}
    />
  );
  const handle = box.current;
  if (!handle) throw new Error("the ref was not attached");
  return { handle, unmount };
}

function panel(): HTMLElement | null {
  return host.querySelector('[role="dialog"][aria-label="Link"]');
}

function field(): HTMLInputElement {
  const found = panel()?.querySelector('input[aria-label="Address"]');
  if (!(found instanceof HTMLInputElement)) {
    throw new Error("the panel has no address field");
  }
  return found;
}

/** A button of the panel, or of the toolbar, by the text or the name on it */
function button(text: string): HTMLButtonElement {
  const found = Array.from(host.querySelectorAll("button")).find(
    (element) =>
      element.textContent === text ||
      element.getAttribute("aria-label") === text
  );
  if (!(found instanceof HTMLButtonElement)) {
    throw new Error(`button not found: ${text}`);
  }
  return found;
}

/** Puts the selection over the first stretch of this text */
function selectText(handle: DocxEditorHandle, needle: string): void {
  const { view } = handle;
  let at = -1;
  view.state.doc.descendants((node, pos) => {
    if (at >= 0 || !node.isText) return true;
    const found = node.text?.indexOf(needle) ?? -1;
    if (found >= 0) at = pos + found;
    return true;
  });
  if (at < 0) throw new Error(`text not found: ${needle}`);
  act(() => {
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, at, at + needle.length)
      )
    );
  });
}

/** The caret one character into the first stretch of this text */
function caretIn(handle: DocxEditorHandle, needle: string): void {
  selectText(handle, needle);
  const { view } = handle;
  act(() => {
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, view.state.selection.from + 1)
      )
    );
  });
}

/** Presses the link key on the editing surface the way a browser would */
function pressLinkKey(handle: DocxEditorHandle): boolean {
  const event = new KeyboardEvent("keydown", {
    key: "k",
    ctrlKey: true,
    bubbles: true,
    cancelable: true,
  });
  act(() => {
    handle.view.dom.dispatchEvent(event);
  });
  return event.defaultPrevented;
}

function press(key: string, options: KeyboardEventInit = {}): void {
  const target = document.activeElement ?? document.body;
  act(() => {
    target.dispatchEvent(
      new KeyboardEvent("keydown", { key, bubbles: true, ...options })
    );
  });
}

/**
 * Types into the address field.
 * React remembers the value it last wrote and takes a plain assignment for no change at all, so
 * the field is written through the native setter and told about it the way a browser would.
 */
function type(value: string): void {
  const input = field();
  const write = Object.getOwnPropertyDescriptor(
    HTMLInputElement.prototype,
    "value"
  )?.set;
  if (!write) throw new Error("no value setter on HTMLInputElement");
  act(() => {
    write.call(input, value);
    input.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

function click(element: HTMLElement): void {
  act(() => element.click());
}

function documentXml(handle: DocxEditorHandle): string {
  return decode(unzipSync(handle.exportBytes())["word/document.xml"]);
}

/** What the focus is on, by its accessible name or the text written on it */
function focused(): string | null {
  const found = document.activeElement;
  if (!(found instanceof HTMLElement)) return null;
  return found.getAttribute("aria-label") ?? found.textContent;
}

describe("opening the link panel", () => {
  it("opens on the link key over selected text, with the field empty and focused", () => {
    const { handle, unmount } = mount(PLAIN);
    selectText(handle, "our terms");
    expect(pressLinkKey(handle)).toBe(true);

    expect(panel()).not.toBeNull();
    expect(field().value).toBe("");
    expect(focused()).toBe("Address");
    unmount();
  });

  it("opens with the address of the link the caret stands in", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    pressLinkKey(handle);

    expect(field().value).toBe(TERMS);
    unmount();
  });

  it("leaves the key alone where there is nothing to link", () => {
    const { handle, unmount } = mount(PLAIN);
    caretIn(handle, "see our terms");

    expect(pressLinkKey(handle)).toBe(false);
    expect(panel()).toBeNull();
    unmount();
  });

  it("opens the same panel from the toolbar button", () => {
    const { handle, unmount } = mount(PLAIN);
    selectText(handle, "our terms");
    click(button("Link"));

    expect(panel()).not.toBeNull();
    unmount();
  });

  it("draws the toolbar button dead where a link could not go on", () => {
    const { handle, unmount } = mount(PLAIN);
    caretIn(handle, "see our terms");

    expect(button("Link").disabled).toBe(true);
    unmount();
  });
});

describe("the panel's buttons", () => {
  it("applies the address typed and closes", () => {
    const { handle, unmount } = mount(PLAIN);
    selectText(handle, "our terms");
    pressLinkKey(handle);
    type(TERMS);
    click(button("Apply"));

    expect(panel()).toBeNull();
    expect(documentXml(handle)).toContain("<w:hyperlink");
    expect(documentXml(handle)).toContain("our terms");
    unmount();
  });

  it("applies on Enter in the field", () => {
    const { handle, unmount } = mount(PLAIN);
    selectText(handle, "our terms");
    pressLinkKey(handle);
    type(TERMS);
    press("Enter");

    expect(panel()).toBeNull();
    expect(documentXml(handle)).toContain("<w:hyperlink");
    unmount();
  });

  it("offers nothing to apply while the field is empty", () => {
    const { handle, unmount } = mount(PLAIN);
    selectText(handle, "our terms");
    pressLinkKey(handle);

    expect(button("Apply").disabled).toBe(true);
    unmount();
  });

  it("takes the link off when the emptied address is applied", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    pressLinkKey(handle);
    type("");
    click(button("Apply"));

    expect(panel()).toBeNull();
    expect(documentXml(handle)).not.toContain("w:hyperlink");
    expect(handle.view.state.doc.child(0).textContent).toBe("see our terms");
    unmount();
  });

  /** Over plain text an empty address has nothing to apply and nothing to take off */
  it("offers nothing on an empty address where there is no link to take off", () => {
    const { handle, unmount } = mount(PLAIN);
    selectText(handle, "our terms");
    pressLinkKey(handle);

    expect(button("Apply").disabled).toBe(true);
    unmount();
  });
});

describe("the panel on a keyboard", () => {
  /** Apply is not on the walk while the address is the one already there: it has nothing to do */
  it("walks its own live controls with Tab and wraps rather than leaving", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    pressLinkKey(handle);

    expect(focused()).toBe("Address");
    // The address is the one already there, so Apply has nothing to do and the walk holds one stop
    press("Tab");
    expect(focused()).toBe("Address");
    press("Tab", { shiftKey: true });
    expect(focused()).toBe("Address");
    unmount();
  });

  it("puts Apply on the walk once the address typed is one to apply", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    pressLinkKey(handle);
    type("https://example.com/prices");

    field().focus();
    press("Tab");
    expect(focused()).toBe("Apply");
    unmount();
  });

  it("closes on Escape and hands the focus back to the paper", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    pressLinkKey(handle);
    press("Escape");

    expect(panel()).toBeNull();
    expect(document.activeElement).toBe(handle.view.dom);
    unmount();
  });

  it("closes on a click outside it", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    pressLinkKey(handle);
    act(() => {
      document.body.dispatchEvent(
        new MouseEvent("mousedown", { bubbles: true })
      );
    });

    expect(panel()).toBeNull();
    unmount();
  });
});
