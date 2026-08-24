// @vitest-environment jsdom
import { EditorView } from "prosemirror-view";
import { act, type RefObject, StrictMode } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createEditorState } from "../editor/createEditor";
import { docxSchema } from "../schema";
import { editorCssVariables } from "../styles/classNames";
import { usePageLayout } from "./usePageLayout";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

let view: EditorView | null = null;
let root: Root | null = null;
let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  if (root) act(() => root?.unmount());
  root = null;
  view?.destroy();
  view = null;
  host.remove();
});

function editor(): EditorView {
  const doc = docxSchema.nodes.doc.create(null, [
    docxSchema.nodes.paragraph.create({}, [docxSchema.text("first")]),
    docxSchema.nodes.paragraph.create({}, [docxSchema.text("second")]),
  ]);
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  view = new EditorView(mount, { state: createEditorState(doc) });
  return view;
}

/** The composition events `view.composing` follows, which is all the hook reads */
function composition(live: EditorView, open: boolean): void {
  live.dom.dispatchEvent(
    new CompositionEvent(open ? "compositionstart" : "compositionend", {
      bubbles: true,
      data: "",
    })
  );
}

/**
 * A measurement writes the sheet height on the overlay box and nothing else does, so counting the
 * writes counts the measurements
 */
function measurements(layer: HTMLElement) {
  const written = vi.spyOn(layer.style, "setProperty");
  return () =>
    written.mock.calls.filter(
      ([name]) => name === editorCssVariables.sheetHeight
    ).length;
}

interface HostProps {
  view: EditorView;
  layer: RefObject<HTMLElement | null>;
  revision: unknown;
}

function Host({ view: live, layer, revision }: HostProps) {
  usePageLayout({ view: live, layer, enabled: true, revision });
  return null;
}

/**
 * A measurement is taken on an animation frame, and a measurement held back by a composition
 * takes the frame again. Waiting one out is what tells "none was taken" apart from "none has
 * been taken yet"
 */
function frame() {
  return act(async () => {
    await new Promise<void>((resolve) => {
      requestAnimationFrame(() => resolve());
    });
  });
}

/** Waits until the measurements taken come to that many, however many frames it takes */
function untilTaken(taken: () => number, count: number) {
  return act(() => vi.waitFor(() => expect(taken()).toBe(count)));
}

function render(props: HostProps) {
  root = createRoot(host);
  const live = root;
  act(() => live.render(<Host {...props} />));
  return (next: HostProps) => act(() => live.render(<Host {...next} />));
}

/** StrictMode mounts, unmounts, and mounts again, running every effect's cleanup in between */
function renderStrict(props: HostProps) {
  root = createRoot(host);
  const live = root;
  act(() =>
    live.render(
      <StrictMode>
        <Host {...props} />
      </StrictMode>
    )
  );
}

describe("the page measurement", () => {
  it("is taken while nothing is being composed", async () => {
    const live = editor();
    const layer: RefObject<HTMLElement | null> = { current: host };
    const taken = measurements(host);
    render({ view: live, layer, revision: live.state.doc });

    await untilTaken(taken, 1);
  });

  it("waits for the composition to end, and then happens once", async () => {
    const live = editor();
    const layer: RefObject<HTMLElement | null> = { current: host };
    composition(live, true);
    const taken = measurements(host);
    const again = render({ view: live, layer, revision: live.state.doc });
    await frame();

    expect(taken()).toBe(0);

    // Every revision the composition brings is one more measurement asked for
    for (const text of ["a", "b", "c"]) {
      live.dispatch(live.state.tr.insertText(text, 1));
      again({ view: live, layer, revision: live.state.doc });
    }
    await frame();
    expect(taken()).toBe(0);

    composition(live, false);
    await untilTaken(taken, 1);
  });

  it("is still taken after StrictMode's simulated remount", async () => {
    const live = editor();
    const layer: RefObject<HTMLElement | null> = { current: host };
    const taken = measurements(host);
    renderStrict({ view: live, layer, revision: live.state.doc });

    await untilTaken(taken, 1);
  });
});
