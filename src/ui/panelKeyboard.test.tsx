// @vitest-environment jsdom
/**
 * The panels the toolbar opens, driven by a keyboard.
 *
 * What is checked here is the interaction layer the panels share: how a panel is announced
 * (`Popover.tsx`), which control holds the one tab stop (`rovingFocus.ts`), and the keys a menu
 * and a grid answer over it (`useMenuKeyboard.ts`, `useGridKeyboard.ts`).
 * Which of them the pointer and the keyboard part company over is the focus: a click on a toolbar
 * button never takes it off the paper, so only a panel the keyboard opened moves it.
 */

import { unzipSync } from "fflate";
import { TextSelection } from "prosemirror-state";
import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { decode, makeDocx } from "../__testing__/docx";
import { renderInto } from "../__testing__/react";
import { DocxEditor, type DocxEditorHandle } from "../DocxEditor";
import { lockSelection } from "../editor/commands/lockCommands";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

const DOCUMENT = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>'
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

function mount() {
  const box: { current: DocxEditorHandle | null } = { current: null };
  const unmount = render(
    <DocxEditor document={DOCUMENT} ref={box} renderImportError={() => null} />
  );
  const handle = box.current;
  if (!handle) throw new Error("the ref was not attached");
  return { handle, unmount };
}

function button(label: string): HTMLButtonElement {
  const found = host.querySelector(`button[aria-label="${label}"]`);
  if (!(found instanceof HTMLButtonElement)) {
    throw new Error(`button not found: ${label}`);
  }
  return found;
}

/** A toolbar control by its name, whether it is a button or one of the selects */
function control(label: string): HTMLElement {
  const found = host.querySelector(`[role="toolbar"] [aria-label="${label}"]`);
  if (!(found instanceof HTMLElement)) {
    throw new Error(`control not found: ${label}`);
  }
  return found;
}

/**
 * Opens a panel the way a keyboard does: the button is focused, and the key lands on it.
 * `click()` raises a click of no detail, which is what a key press on a button reports.
 */
function openWithKeyboard(label: string): HTMLButtonElement {
  const trigger = button(label);
  trigger.focus();
  act(() => trigger.click());
  return trigger;
}

/** Opens with a pointer click, which leaves the test's existing focus unchanged. */
function openWithPointer(label: string): void {
  const trigger = button(label);
  act(() => {
    trigger.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    trigger.dispatchEvent(
      new MouseEvent("click", { bubbles: true, detail: 1 })
    );
  });
}

function press(key: string, options: KeyboardEventInit = {}): void {
  const target = document.activeElement ?? document.body;
  act(() => {
    target.dispatchEvent(
      new KeyboardEvent("keydown", { key, bubbles: true, ...options })
    );
  });
}

function panel(): HTMLElement | null {
  return host.querySelector(".docx-editor-popover");
}

/** What the focus is on, by its accessible name or the text written on it */
function focused(): string | null {
  const found = document.activeElement;
  if (!(found instanceof HTMLElement)) return null;
  return found.getAttribute("aria-label") ?? found.textContent;
}

/** Every control of the open panel that Tab can reach. One of them, by the roving stop */
function tabbable(): string[] {
  return Array.from(panel()?.querySelectorAll("button") ?? [])
    .filter((control) => control.getAttribute("tabindex") === "0")
    .map(
      (control) =>
        control.getAttribute("aria-label") ?? control.textContent ?? ""
    );
}

/** Every control of the toolbar that Tab can reach. One of them, by the roving stop */
function tabbableControls(): string[] {
  return Array.from(
    host.querySelectorAll<HTMLElement>(
      '[role="toolbar"] .docx-editor-toolbar-btn, [role="toolbar"] .docx-editor-toolbar-select'
    )
  )
    .filter((control) => control.getAttribute("tabindex") === "0")
    .map((control) => control.getAttribute("aria-label") ?? "");
}

/** The one control of the toolbar the keyboard starts on: nothing has been edited, so undo and redo are dead */
const FIRST_CONTROL = "Zoom";

describe("the toolbar as one tab stop", () => {
  it("leaves a single control reachable by Tab, and it is one that has something to do", () => {
    const { unmount } = mount();

    // The history buttons are drawn first, and there is nothing yet to undo or to redo
    expect(button("Undo").disabled).toBe(true);
    expect(button("Redo").disabled).toBe(true);
    expect(tabbableControls()).toEqual([FIRST_CONTROL]);
    unmount();
  });

  it("walks the controls sideways, wrapping at the ends", () => {
    const { unmount } = mount();
    control(FIRST_CONTROL).focus();

    press("ArrowRight");
    expect(focused()).toBe("Style");
    press("ArrowLeft");
    expect(focused()).toBe(FIRST_CONTROL);
    // It is the first control with anything to do, so the walk wraps round to the far end
    press("ArrowLeft");
    expect(focused()).toBe("Show comments");
    unmount();
  });

  it("jumps to the ends with Home and End", () => {
    const { unmount } = mount();
    control(FIRST_CONTROL).focus();

    press("End");
    expect(focused()).toBe("Show comments");
    press("Home");
    expect(focused()).toBe(FIRST_CONTROL);
    unmount();
  });

  it("carries the stop with the control the walk left off on", () => {
    const { unmount } = mount();
    control(FIRST_CONTROL).focus();

    press("ArrowRight");
    press("ArrowRight");
    press("ArrowRight");
    expect(focused()).toBe("Font size");
    expect(tabbableControls()).toEqual(["Font size"]);
    unmount();
  });

  it("hands the stop on when the control holding it goes dead", () => {
    const { handle, unmount } = mount();
    control(FIRST_CONTROL).focus();
    press("ArrowRight");
    press("ArrowRight");
    press("ArrowRight");
    press("ArrowRight");
    expect(tabbableControls()).toEqual(["Bold"]);

    // Nothing can be written on text a content control has shut
    act(() => {
      const { view } = handle;
      const paragraph = view.state.doc.child(0);
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(view.state.doc, 1, 1 + paragraph.content.size)
        )
      );
      lockSelection(view.state, (tr) => view.dispatch(tr));
    });

    expect(button("Bold").disabled).toBe(true);
    const held = tabbableControls();
    expect(held).toHaveLength(1);
    expect(held).not.toContain("Bold");
    unmount();
  });
});

describe("what a panel says it is", () => {
  it("names the panel a button opens on the button itself", () => {
    const { unmount } = mount();
    expect(button("Text color").getAttribute("aria-haspopup")).toBe("dialog");
    expect(button("Alignment").getAttribute("aria-haspopup")).toBe("menu");
    unmount();
  });

  it("announces a panel of controls as a dialog named after its button", () => {
    const { unmount } = mount();
    openWithPointer("Text color");

    expect(panel()?.getAttribute("role")).toBe("dialog");
    expect(panel()?.getAttribute("aria-label")).toBe("Text color");
    unmount();
  });

  it("leaves a panel holding a menu to the menu, so neither role is doubled", () => {
    const { unmount } = mount();
    openWithPointer("Alignment");

    expect(panel()?.getAttribute("role")).toBeNull();
    expect(
      panel()?.querySelector('[role="menu"]')?.getAttribute("aria-label")
    ).toBe("Paragraph alignment");
    unmount();
  });

  it("draws the palette as a grid of rows with the color in force selected", () => {
    const { unmount } = mount();
    openWithPointer("Text color");
    const grid = panel()?.querySelector('[role="grid"]');

    expect(grid?.getAttribute("aria-label")).toBe("Colors");
    // The eight rows of the palette, and the row holding the entry that withdraws the color
    expect(grid?.querySelectorAll('[role="row"]')).toHaveLength(9);
    expect(grid?.querySelectorAll('[role="gridcell"]')).toHaveLength(81);
    // Nothing is painted yet, so what is in force is the withdrawal
    expect(grid?.querySelector('[aria-selected="true"]')?.textContent).toBe(
      "None"
    );
    unmount();
  });

  it("speaks the size the table grid is standing on", () => {
    const { unmount } = mount();
    openWithPointer("Insert table");
    const readout = panel()?.parentElement?.querySelector("p");

    expect(readout?.getAttribute("aria-live")).toBe("polite");
    expect(readout?.textContent).toBe("Pick a size");
    unmount();
  });
});

describe("a panel the keyboard opened", () => {
  it("moves the focus onto the row that carries the current choice", () => {
    const { unmount } = mount();
    openWithKeyboard("Alignment");

    expect(focused()).toBe("Align left");
    expect(tabbable()).toEqual(["Align left"]);
    unmount();
  });

  it("walks the rows with the arrows and applies the one chosen", () => {
    const { handle, unmount } = mount();
    openWithKeyboard("Alignment");

    press("ArrowDown");
    expect(focused()).toBe("Align center");
    act(() => {
      const row = document.activeElement;
      if (row instanceof HTMLElement) row.click();
    });

    expect(handle.view.state.doc.child(0).attrs.pPr).toContain(
      '<w:jc w:val="center"/>'
    );
    unmount();
  });

  it("hands the focus back to the button on Escape", () => {
    const { unmount } = mount();
    const trigger = openWithKeyboard("Alignment");
    expect(focused()).toBe("Align left");

    press("Escape");
    expect(panel()).toBeNull();
    expect(document.activeElement).toBe(trigger);
    unmount();
  });

  it("hands the focus back to the button once a choice is made", () => {
    const { unmount } = mount();
    const trigger = openWithKeyboard("Alignment");

    act(() => {
      const row = document.activeElement;
      if (row instanceof HTMLElement) row.click();
    });
    expect(panel()).toBeNull();
    expect(document.activeElement).toBe(trigger);
    unmount();
  });

  it("closes on Tab rather than walking out into the rows behind it", () => {
    const { unmount } = mount();
    const trigger = openWithKeyboard("Alignment");

    press("Tab");
    expect(panel()).toBeNull();
    expect(document.activeElement).toBe(trigger);
    unmount();
  });

  it("swallows a key that would only scroll the page out from under the panel", () => {
    const { unmount } = mount();
    openWithKeyboard("Alignment");

    press("PageDown");
    expect(panel()).not.toBeNull();
    expect(focused()).toBe("Align left");
    unmount();
  });
});

describe("a panel the pointer opened", () => {
  it("leaves the focus on the paper, where the selection it is about to act on is drawn", () => {
    const { handle, unmount } = mount();
    act(() => handle.view.focus());
    openWithPointer("Alignment");

    expect(panel()).not.toBeNull();
    expect(document.activeElement).toBe(handle.view.dom);
    unmount();
  });

  it("leaves it there when it closes too", () => {
    const { handle, unmount } = mount();
    act(() => handle.view.focus());
    openWithPointer("Alignment");

    press("Escape");
    expect(panel()).toBeNull();
    expect(document.activeElement).toBe(handle.view.dom);
    unmount();
  });
});

describe("the color grid on a keyboard", () => {
  it("opens on the color in force and walks the swatches in both directions", () => {
    const { unmount } = mount();
    openWithKeyboard("Text color");
    // Nothing is painted, so the walk starts at the entry that withdraws the color
    expect(focused()).toBe("None");

    press("ArrowDown");
    expect(focused()).toBe("#000000");
    press("ArrowRight");
    expect(focused()).toBe("#434343");
    press("ArrowDown");
    expect(focused()).toBe("#ff0000");
    press("ArrowUp");
    expect(focused()).toBe("#434343");
    unmount();
  });

  it("stops at the edges rather than wrapping into the next row", () => {
    const { unmount } = mount();
    openWithKeyboard("Text color");

    press("ArrowDown");
    press("ArrowLeft");
    expect(focused()).toBe("#000000");
    press("End");
    expect(focused()).toBe("#ffffff");
    press("ArrowRight");
    expect(focused()).toBe("#ffffff");
    unmount();
  });

  it("jumps to the corners of the whole grid with the modifier held", () => {
    const { unmount } = mount();
    openWithKeyboard("Text color");

    press("End", { ctrlKey: true });
    expect(focused()).toBe("#4c1130");
    press("Home", { ctrlKey: true });
    expect(focused()).toBe("None");
    unmount();
  });

  it("paints the swatch the keyboard walked to and hands the focus back", () => {
    const { handle, unmount } = mount();
    act(() => {
      const { view } = handle;
      const paragraph = view.state.doc.child(0);
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(view.state.doc, 1, 1 + paragraph.content.size)
        )
      );
    });
    const trigger = openWithKeyboard("Text color");

    press("ArrowDown");
    press("ArrowRight");
    act(() => {
      const cell = document.activeElement;
      if (cell instanceof HTMLElement) cell.click();
    });

    expect(panel()).toBeNull();
    expect(document.activeElement).toBe(trigger);
    expect(
      decode(unzipSync(handle.exportBytes())["word/document.xml"])
    ).toContain('<w:color w:val="434343"/>');
    unmount();
  });

  it("stages a system-picker color until the apply button is pressed", () => {
    const { handle, unmount } = mount();
    act(() => {
      const { view } = handle;
      const paragraph = view.state.doc.child(0);
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(view.state.doc, 1, 1 + paragraph.content.size)
        )
      );
    });
    openWithPointer("Text color");
    const picker = panel()?.querySelector<HTMLInputElement>(
      'input[aria-label="Choose custom color"]'
    );
    const apply = panel()?.querySelector<HTMLButtonElement>(
      'button[aria-label="Apply color"]'
    );
    if (!picker || !apply) throw new Error("custom color controls not found");
    const write = Object.getOwnPropertyDescriptor(
      HTMLInputElement.prototype,
      "value"
    )?.set;
    if (!write) throw new Error("no value setter on HTMLInputElement");

    act(() => {
      write.call(picker, "#abcdef");
      picker.dispatchEvent(new Event("input", { bubbles: true }));
    });

    expect(panel()).not.toBeNull();
    expect(
      decode(unzipSync(handle.exportBytes())["word/document.xml"])
    ).not.toContain("<w:color");
    expect(
      panel()?.querySelector<HTMLInputElement>(
        'input[aria-label="Custom color"]'
      )?.value
    ).toBe("#ABCDEF");
    act(() => apply.click());
    expect(
      decode(unzipSync(handle.exportBytes())["word/document.xml"])
    ).toContain('<w:color w:val="ABCDEF"/>');
    unmount();
  });

  it("accepts an exact custom color and stores normalized OOXML RGB", () => {
    const { handle, unmount } = mount();
    act(() => {
      const { view } = handle;
      const paragraph = view.state.doc.child(0);
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(view.state.doc, 1, 1 + paragraph.content.size)
        )
      );
    });
    openWithPointer("Text color");
    const field = panel()?.querySelector<HTMLInputElement>(
      'input[aria-label="Custom color"]'
    );
    const apply = panel()?.querySelector<HTMLButtonElement>(
      'button[type="submit"]'
    );
    if (!field || !apply) throw new Error("custom color controls not found");
    const write = Object.getOwnPropertyDescriptor(
      HTMLInputElement.prototype,
      "value"
    )?.set;
    if (!write) throw new Error("no value setter on HTMLInputElement");

    act(() => {
      write.call(field, "1a2b3c");
      field.dispatchEvent(new Event("input", { bubbles: true }));
    });
    act(() => apply.click());

    expect(panel()).toBeNull();
    expect(
      decode(unzipSync(handle.exportBytes())["word/document.xml"])
    ).toContain('<w:color w:val="1A2B3C"/>');
    unmount();
  });

  it("moves from the palette into the custom controls with Tab", () => {
    const { unmount } = mount();
    openWithKeyboard("Text color");

    press("Tab");
    expect(focused()).toBe("Choose custom color");
    expect(panel()).not.toBeNull();
    press("Tab", { shiftKey: true });
    expect(focused()).toBe("None");
    unmount();
  });
});
