// @vitest-environment jsdom
/**
 * The tooltips on the toolbar controls.
 *
 * How long a tooltip waits and how it looks belong to the CSS, and jsdom draws nothing, so
 * what is checked here is everything short of the drawing: that every control carries its
 * own accessible name as its tooltip text, that nothing is left relying on the browser's
 * own `title`, and that the CSS reads the same attribute the components write and waits
 * the 300ms it is supposed to.
 * That the box actually fades in after that wait, below the control and over the paper,
 * has to be seen in a browser.
 */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { EDITING } from "../__testing__/mode";
import { renderInto } from "../__testing__/react";
import { DocxEditor } from "../DocxEditor";
import { editorAttributes } from "../styles/classNames";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

const TOOLTIP = editorAttributes.tooltip;

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
    <DocxEditor
      document={DOCUMENT}
      mode={EDITING}
      renderImportError={() => null}
    />
  );
}

function toolbar(): HTMLElement {
  const found = host.querySelector('[role="toolbar"]');
  if (!(found instanceof HTMLElement))
    throw new Error("the toolbar is missing");
  return found;
}

/** Everything inside the toolbar that shows a tooltip, including any open panel */
function tooltipTargets(): HTMLElement[] {
  return Array.from(toolbar().querySelectorAll<HTMLElement>(`[${TOOLTIP}]`));
}

/**
 * The name a screen reader reads for the control a tooltip belongs to. It is the target
 * itself, except where the target is the box wrapped around a control that cannot draw a
 * tooltip of its own.
 */
function accessibleName(target: HTMLElement): string | null {
  const own = target.getAttribute("aria-label");
  if (own !== null) return own;
  return (
    target.querySelector("[aria-label]")?.getAttribute("aria-label") ?? null
  );
}

function tooltipOf(label: string): string | null {
  const found = tooltipTargets().find(
    (target) => accessibleName(target) === label
  );
  if (found === undefined) throw new Error(`no control named: ${label}`);
  return found.getAttribute(TOOLTIP);
}

function click(label: string): void {
  const found = toolbar().querySelector(`button[aria-label="${label}"]`);
  if (!(found instanceof HTMLButtonElement)) {
    throw new Error(`button not found: ${label}`);
  }
  act(() => found.click());
}

describe("the toolbar tooltips", () => {
  it("say exactly what the control is named, so the two cannot drift apart", () => {
    const unmount = mount();
    const targets = tooltipTargets();

    expect(targets.length).toBeGreaterThan(0);
    for (const target of targets) {
      expect(target.getAttribute(TOOLTIP)).toBe(accessibleName(target));
    }
    unmount();
  });

  it("cover the buttons, the pickers, and the panels they open", () => {
    const unmount = mount();

    expect(tooltipOf("Bold")).toBe("Bold");
    // The button that opens a panel is a toolbar button like any other
    expect(tooltipOf("Text color")).toBe("Text color");
    // A select cannot draw one, so the box around it carries the text
    expect(tooltipOf("Font")).toBe("Font");
    expect(tooltipOf("Font size")).toBe("Font size");
    unmount();
  });

  it("keeps palette labels accessible without tooltip overflow", () => {
    const unmount = mount();
    click("Text color");

    const swatch = toolbar().querySelector('button[aria-label="#ff0000"]');
    expect(swatch).not.toBeNull();
    expect(swatch?.hasAttribute(TOOLTIP)).toBe(false);
    unmount();
  });

  it("are taken away by Escape and brought back by the next move", () => {
    const unmount = mount();
    const HUSHED = editorAttributes.tooltipsHushed;
    expect(toolbar().hasAttribute(HUSHED)).toBe(false);

    act(() => {
      document.dispatchEvent(
        new KeyboardEvent("keydown", { key: "Escape", bubbles: true })
      );
    });
    expect(toolbar().hasAttribute(HUSHED)).toBe(true);

    act(() => {
      document.dispatchEvent(new Event("pointermove", { bubbles: true }));
    });
    expect(toolbar().hasAttribute(HUSHED)).toBe(false);
    unmount();
  });

  it("leave no control waiting on the browser's own title", () => {
    const unmount = mount();
    expect(toolbar().querySelectorAll("[title]")).toHaveLength(0);

    // The palette is only rendered while the panel is open, swatches and all
    click("Text color");
    expect(toolbar().querySelectorAll("[title]")).toHaveLength(0);
    unmount();
  });
});

/**
 * The wait is the whole point of dropping the native `title`, and it is written in one
 * place only: the CSS.
 */
describe("the tooltip rules in editor.css", () => {
  const css = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), "../styles/editor.css"),
    "utf8"
  );

  /** What one rule declares, found by the selector it starts with */
  function declarations(selector: string): string {
    const found = css.match(
      new RegExp(`\\[${TOOLTIP}\\]${selector}[\\s\\S]*?\\{([\\s\\S]*?)\\}`)
    )?.[1];
    if (found === undefined) throw new Error(`rule not found: ${selector}`);
    return found;
  }

  it("draws the text from the same attribute the components write", () => {
    expect(css).toContain(`content: attr(${TOOLTIP})`);
  });

  it("shows nothing until the pointer has rested on the control for 300ms", () => {
    const shown = declarations(":hover::after");
    expect(shown).toContain("opacity: 1");
    expect(shown).toMatch(/transition:[\s\S]*300ms/);
  });

  it("waits the same 300ms for keyboard focus, and ignores focus taken by a click", () => {
    expect(css).toContain(`[${TOOLTIP}]:focus-visible::after`);
    // A select is focused inside the box that carries its tooltip
    expect(css).toContain(`[${TOOLTIP}]:has(:focus-visible)::after`);
  });

  it("draws none at all while the box they are inside is hushed", () => {
    // The rule has to come after the one that shows a tooltip, since the two weigh the same
    const hush = css.indexOf(
      `[${editorAttributes.tooltipsHushed}] [${TOOLTIP}]::after`
    );
    expect(hush).toBeGreaterThan(css.indexOf(`[${TOOLTIP}]:hover::after`));
    expect(css.slice(hush, hush + 120)).toContain("content: none");
  });

  it("carries no delay in the hidden state, so leaving the control is immediate", () => {
    const hidden = declarations("::after");
    expect(hidden).toContain("opacity: 0");
    expect(hidden).toContain("visibility: hidden");
    expect(hidden).not.toContain("300ms");
  });
});
