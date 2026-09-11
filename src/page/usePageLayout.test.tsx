// @vitest-environment jsdom
import { EditorView } from "prosemirror-view";
import { act, type RefObject, StrictMode } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createEditorState } from "../editor/createEditor";
import { docxSchema } from "../schema";
import { editorCssVariables } from "../styles/classNames";
import type { DemandBand } from "./demands";
import { FOOTNOTE_BAND } from "./demands/footnoteDemands";
import { A4_PAGE_PIXELS, type SectionPixels } from "./pageLayout";
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

/** The sheet heights written on the overlay box, in the order they were measured */
function sheetHeights(layer: HTMLElement) {
  const written = vi.spyOn(layer.style, "setProperty");
  return () =>
    written.mock.calls
      .filter(([name]) => name === editorCssVariables.sheetHeight)
      .map(([, value]) => value);
}

interface HostProps {
  view: EditorView;
  layer: RefObject<HTMLElement | null>;
  revision: unknown;
  sections?: readonly SectionPixels[];
  bands?: ReadonlyMap<string, DemandBand>;
}

function Host({ view: live, layer, revision, sections, bands }: HostProps) {
  usePageLayout({
    view: live,
    layer,
    enabled: true,
    revision,
    sections,
    bands,
  });
  return null;
}

/** The footnote band with one footnote of this height in it */
function footnoteBand(height: number): ReadonlyMap<string, DemandBand> {
  return new Map([
    [
      FOOTNOTE_BAND,
      { overhead: 16, heights: new Map([["footnote:1", height]]) },
    ],
  ]);
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

  it("stretches the sheet to the bottom margin of the page it ends on", async () => {
    const live = editor();
    // The padding the sheet is drawn with is the first section's (`editor/createEditor`), and
    // the document ends on a section whose bottom margin is deeper than that
    live.dom.style.padding = "20px";
    const sections: readonly SectionPixels[] = [
      {
        untilPos: Number.POSITIVE_INFINITY,
        pixels: { ...A4_PAGE_PIXELS, bodyHeight: 500, marginBottom: 80 },
        type: null,
      },
    ];
    const layer: RefObject<HTMLElement | null> = { current: host };
    const heights = sheetHeights(host);
    render({ view: live, layer, revision: live.state.doc, sections });

    // The sheet's own top padding, one page of body, and the margin under that page
    await act(() =>
      vi.waitFor(() => expect(heights()).toEqual([`${20 + 500 + 80}px`]))
    );
  });

  it("lays the pages out again when a footnote grows", async () => {
    const live = editor();
    const layer: RefObject<HTMLElement | null> = { current: host };
    const taken = measurements(host);
    const bands = footnoteBand(40);
    const again = render({
      view: live,
      layer,
      revision: live.state.doc,
      bands,
    });
    await untilTaken(taken, 1);

    // The same heights handed in again ask for nothing
    again({ view: live, layer, revision: live.state.doc, bands });
    await frame();
    expect(taken()).toBe(1);

    again({
      view: live,
      layer,
      revision: live.state.doc,
      bands: footnoteBand(64),
    });
    await untilTaken(taken, 2);
  });

  it("is still taken after StrictMode's simulated remount", async () => {
    const live = editor();
    const layer: RefObject<HTMLElement | null> = { current: host };
    const taken = measurements(host);
    renderStrict({ view: live, layer, revision: live.state.doc });

    await untilTaken(taken, 1);
  });
});
