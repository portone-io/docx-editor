// @vitest-environment jsdom
/**
 * The card that stands by a link, driven the way a person drives it: the caret moving into a link,
 * a click on what the card offers, and the keys that take it away.
 *
 * Which link the selection sits inside is `editor/commands/linkCommands.test.ts`; what is checked
 * here is that the card appears for it, says what it points at, does not take focus as it appears,
 * and offers only what the commands behind it would go through with.
 */

import { unzipSync } from "fflate";
import { TextSelection } from "prosemirror-state";
import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { decode, makeDocx, makeLinkedDocx } from "../__testing__/docx";
import { renderInto } from "../__testing__/react";
import {
  DocxEditor,
  type DocxEditorHandle,
  type DocxEditorMode,
} from "../DocxEditor";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

const TERMS = "https://example.com/terms";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const lockedControl = (inner: string) =>
  '<w:sdt><w:sdtPr><w:id w:val="7"/>' +
  '<w:lock w:val="sdtContentLocked"/></w:sdtPr>' +
  `<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const linkTo = (text: string) =>
  `<w:hyperlink r:id="rId9">${run(text)}</w:hyperlink>`;

const LINKED = makeLinkedDocx(
  `<w:p>${run("see ")}${linkTo("our terms")}</w:p>`,
  {
    rId9: TERMS,
  }
);

/** A link naming a bookmark, which is a link with no address to show */
const ANCHORED = makeDocx(
  `<w:p>${run("see ")}<w:hyperlink w:anchor="chapter3">` +
    `${run("chapter three")}</w:hyperlink></w:p>`
);

/** The same link, inside a control that says its contents may not be edited */
const LOCKED = makeLinkedDocx(
  `<w:p>${run("see ")}${lockedControl(linkTo("our terms"))}</w:p>`,
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

function mount(bytes: Uint8Array, mode?: DocxEditorMode) {
  const box: { current: DocxEditorHandle | null } = { current: null };
  const unmount = render(
    <DocxEditor
      document={bytes}
      ref={box}
      mode={mode}
      renderImportError={() => null}
    />
  );
  const handle = box.current;
  if (!handle) throw new Error("the ref was not attached");
  return { handle, unmount };
}

function card(): HTMLElement | null {
  return host.querySelector('[role="group"][aria-label="Link"]');
}

function panel(): HTMLElement | null {
  return host.querySelector('[role="dialog"][aria-label="Link"]');
}

/** The text the card shows before its buttons, which is the address or the note in place of one */
function shown(): string {
  const first = card()?.firstElementChild;
  if (!(first instanceof HTMLElement)) throw new Error("the card is empty");
  return first.textContent ?? "";
}

// The buttons carry icons alone, so their accessible name is the label
function cardButtons(): string[] {
  return Array.from(card()?.querySelectorAll("button") ?? []).map(
    (control) => control.getAttribute("aria-label") ?? ""
  );
}

function cardButton(name: string): HTMLButtonElement {
  const found = Array.from(card()?.querySelectorAll("button") ?? []).find(
    (control) => control.getAttribute("aria-label") === name
  );
  if (!(found instanceof HTMLButtonElement)) {
    throw new Error(`the card has no button: ${name}`);
  }
  return found;
}

/** The stretch the first occurrence of this text covers */
function rangeOf(handle: DocxEditorHandle, needle: string) {
  let at = -1;
  handle.view.state.doc.descendants((node, pos) => {
    if (at >= 0 || !node.isText) return true;
    const found = node.text?.indexOf(needle) ?? -1;
    if (found >= 0) at = pos + found;
    return true;
  });
  if (at < 0) throw new Error(`text not found: ${needle}`);
  return { from: at, to: at + needle.length };
}

function selectRange(handle: DocxEditorHandle, from: number, to = from): void {
  const { view } = handle;
  act(() => {
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to))
    );
  });
}

/** The whole of the first stretch of this text selected */
function selectText(handle: DocxEditorHandle, needle: string): void {
  const { from, to } = rangeOf(handle, needle);
  selectRange(handle, from, to);
}

/** The caret one character into the first stretch of this text */
function caretIn(handle: DocxEditorHandle, needle: string): void {
  selectRange(handle, rangeOf(handle, needle).from + 1);
}

/** The composition events `view.composing` follows */
function startComposition(handle: DocxEditorHandle): void {
  act(() => {
    handle.view.dom.dispatchEvent(
      new CompositionEvent("compositionstart", { bubbles: true, data: "" })
    );
  });
}

function press(key: string): void {
  const target = document.activeElement ?? document.body;
  act(() => {
    target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
  });
}

function click(element: HTMLElement): void {
  act(() => element.click());
}

function documentXml(handle: DocxEditorHandle): string {
  return decode(unzipSync(handle.exportBytes())["word/document.xml"]);
}

describe("the card that says where a link points", () => {
  it("appears with the caret in a link, showing the address", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");

    expect(card()).not.toBeNull();
    expect(shown()).toBe(TERMS);
    expect(cardButtons()).toEqual(["Open", "Edit", "Remove"]);
    unmount();
  });

  it("stays for a selection that does not leave the link", () => {
    const { handle, unmount } = mount(LINKED);
    selectText(handle, "our terms");

    expect(shown()).toBe(TERMS);
    unmount();
  });

  it("is absent with the caret outside every link", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "see ");

    expect(card()).toBeNull();
    unmount();
  });

  it("is absent for a selection that runs out of the link", () => {
    const { handle, unmount } = mount(LINKED);
    const { to } = rangeOf(handle, "our terms");
    selectRange(handle, to - 2, to + 1);

    expect(card()).toBeNull();
    unmount();
  });

  it("is absent while an IME is composing", () => {
    const { handle, unmount } = mount(LINKED);
    startComposition(handle);
    caretIn(handle, "our terms");

    expect(handle.view.composing).toBe(true);
    expect(card()).toBeNull();
    unmount();
  });

  it("says a link naming a place in the document has no address", () => {
    const { handle, unmount } = mount(ANCHORED);
    caretIn(handle, "chapter three");

    expect(shown()).toBe(
      "This link points at a place in the document rather than at an address."
    );
    expect(cardButtons()).toEqual(["Edit", "Remove"]);
    unmount();
  });

  it("goes away on Escape", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    press("Escape");

    expect(card()).toBeNull();
    unmount();
  });
});

describe("the card's opening of an address", () => {
  /** The same schemes `ui/openAddress` follows: an address it will not open is offered greyed */
  it("offers to open an http address and not one with a scheme we will not follow", () => {
    const risky = makeLinkedDocx(`<w:p>${linkTo("run me")}</w:p>`, {
      rId9: "javascript:alert(1)",
    });
    const there = mount(risky);
    caretIn(there.handle, "run me");
    expect(cardButton("Open").disabled).toBe(true);
    there.unmount();

    const fine = mount(LINKED);
    caretIn(fine.handle, "our terms");
    expect(cardButton("Open").disabled).toBe(false);
    fine.unmount();
  });
});

describe("the card while the text under it is edited", () => {
  /**
   * The card is anchored to the link's own stretch rather than to the caret, so a keystroke inside
   * the link leaves the very same box standing: a card that remounted per keystroke would flicker.
   */
  it("keeps the same box while text is typed inside the link", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    const before = card();
    const { view } = handle;
    act(() => {
      view.dispatch(view.state.tr.insertText("x", view.state.selection.from));
    });

    expect(card()).toBe(before);
    expect(shown()).toBe(TERMS);
    unmount();
  });

  it("takes the link off, and goes with it", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    click(cardButton("Remove"));

    expect(card()).toBeNull();
    expect(documentXml(handle)).not.toContain("w:hyperlink");
    expect(handle.view.state.doc.child(0).textContent).toBe("see our terms");
    unmount();
  });

  it("opens the link panel with the address already in it", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    click(cardButton("Edit"));

    // The panel is the same one Cmd+K opens, and it stands in place of the card
    expect(panel()?.querySelector("input")).toHaveProperty("value", TERMS);
    expect(card()).toBeNull();
    unmount();
  });
});

describe("the card and the focus", () => {
  /** Typing has to be able to continue while the card stands, so it takes no focus on appearance */
  it("does not take focus when it appears", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");

    expect(card()?.contains(document.activeElement)).toBe(false);
    unmount();
  });

  it("returns focus to the paper when Escape closes it from a button", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    cardButton("Open").focus();

    expect(document.activeElement).toBe(cardButton("Open"));
    press("Escape");

    expect(card()).toBeNull();
    expect(document.activeElement).toBe(handle.view.dom);
    unmount();
  });

  it("refuses the mousedown its buttons are pressed with, which is what keeps the selection", () => {
    const { handle, unmount } = mount(LINKED);
    caretIn(handle, "our terms");
    const event = new MouseEvent("mousedown", {
      bubbles: true,
      cancelable: true,
    });
    act(() => {
      cardButton("Open").dispatchEvent(event);
    });

    expect(event.defaultPrevented).toBe(true);
    unmount();
  });
});

describe("the card where nothing may be edited", () => {
  it("offers a reader the address and the opening of it, and nothing else", () => {
    const { handle, unmount } = mount(LINKED, { kind: "readOnly" });
    caretIn(handle, "our terms");

    expect(shown()).toBe(TERMS);
    expect(cardButtons()).toEqual(["Open"]);
    unmount();
  });

  it("shows the address of a link inside a settled part and offers no editing", () => {
    const { handle, unmount } = mount(LOCKED);
    caretIn(handle, "our terms");

    expect(shown()).toBe(TERMS);
    expect(cardButtons()).toEqual(["Open"]);
    unmount();
  });
});
