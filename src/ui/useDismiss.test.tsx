// @vitest-environment jsdom
/**
 * The one rule that closes an open panel.
 *
 * The rule is driven here through a real toolbar panel rather than through a component built for
 * the test, since what it has to get right is which event counts as a move of the screen under
 * the panel and which does not.
 */

import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { renderInto } from "../__testing__/react";
import { DocxEditor } from "../DocxEditor";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

const DOCUMENT = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>'
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
  return render(
    <DocxEditor document={DOCUMENT} renderImportError={() => null} />
  );
}

/** Opens the alignment panel, which holds a menu of four rows */
function openPanel(): void {
  const found = host.querySelector('button[aria-label="Alignment"]');
  if (!(found instanceof HTMLButtonElement)) {
    throw new Error("the alignment button is missing");
  }
  act(() => found.click());
}

function panel(): HTMLElement | null {
  return host.querySelector('[role="menu"][aria-label="Paragraph alignment"]');
}

function firstRow(): HTMLElement {
  const found = panel()?.querySelector('[role="menuitemradio"]');
  if (!(found instanceof HTMLElement)) throw new Error("the panel is not open");
  return found;
}

function scrollOn(target: EventTarget): void {
  act(() => {
    target.dispatchEvent(new Event("scroll"));
  });
}

describe("an open toolbar panel", () => {
  it("is left alone by a scroll that started inside it", () => {
    const unmount = mount();
    openPanel();
    expect(panel()).not.toBeNull();

    scrollOn(firstRow());
    expect(panel()).not.toBeNull();
    unmount();
  });

  it("closes on a scroll anywhere else, which leaves its placement stale", () => {
    const unmount = mount();
    openPanel();
    expect(panel()).not.toBeNull();

    scrollOn(document);
    expect(panel()).toBeNull();
    unmount();
  });

  it("closes on Escape", () => {
    const unmount = mount();
    openPanel();

    act(() => {
      document.dispatchEvent(
        new KeyboardEvent("keydown", { key: "Escape", bubbles: true })
      );
    });
    expect(panel()).toBeNull();
    unmount();
  });

  it("closes on a click outside it and stands through one inside", () => {
    const unmount = mount();
    openPanel();

    act(() => {
      firstRow().dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    });
    expect(panel()).not.toBeNull();

    act(() => {
      document.body.dispatchEvent(
        new MouseEvent("mousedown", { bubbles: true })
      );
    });
    expect(panel()).toBeNull();
    unmount();
  });
});
