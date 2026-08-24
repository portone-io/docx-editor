// @vitest-environment jsdom
import { undo } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import {
  type EditorState,
  NodeSelection,
  TextSelection,
} from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import {
  drawingRun,
  inlineDrawingXml,
  makeImageDocx,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { pxToEmu } from "../../ooxml/image";
import { editorClassNames } from "../../styles/classNames";
import { createEditorState, createEditorView } from "../createEditor";
import {
  buildResizeImageTransaction,
  type ImageCorner,
  resizedImagePx,
} from "./imageResize";

/** 200 x 100 px, the size the test drawing records */
const START = { width: 200, height: 100 };

describe("the size a corner drag asks for", () => {
  it("follows the pointer for a drag along the image's own diagonal", () => {
    expect(resizedImagePx(START, "se", 100, 50)).toEqual({
      width: 300,
      height: 150,
    });
    expect(resizedImagePx(START, "nw", -100, -50)).toEqual({
      width: 300,
      height: 150,
    });
  });

  it("shrinks the image when the corner is dragged inwards", () => {
    expect(resizedImagePx(START, "se", -100, -50)).toEqual({
      width: 100,
      height: 50,
    });
  });

  it("resizes from a drag along one axis alone, keeping the ratio", () => {
    const wider = resizedImagePx(START, "se", 100, 0);
    expect(wider).toEqual({ width: 250, height: 125 });
    const taller = resizedImagePx(START, "se", 0, 50);
    expect(taller).toEqual({ width: 250, height: 125 });
  });

  it("reads each corner's outward direction from where it sits", () => {
    const grown = { width: 250, height: 125 };
    const cases: [ImageCorner, number, number][] = [
      ["se", 100, 0],
      ["sw", -100, 0],
      ["ne", 0, -50],
      ["nw", -100, 0],
    ];
    for (const [corner, dx, dy] of cases) {
      expect(resizedImagePx(START, corner, dx, dy)).toEqual(grown);
    }
  });

  it("stops at a size there is still something left to grab of", () => {
    const smallest = resizedImagePx(START, "se", -1000, -1000);
    expect(smallest).toEqual({ width: 32, height: 16 });
    // The ratio holds even at the limit
    expect(smallest.width / smallest.height).toBeCloseTo(2, 6);
  });
});

const SEAL = `<w:p><w:r><w:t xml:space="preserve">Signature</w:t></w:r>${drawingRun(
  inlineDrawingXml({ descr: "Seal" })
)}</w:p>`;

/** Where the one image of the document sits */
function imagePos(doc: PMNode): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (node.type.name === "image") found = pos;
    return true;
  });
  if (found < 0) throw new Error("the document has no image");
  return found;
}

function imageAt(state: EditorState): PMNode {
  const node = state.doc.nodeAt(imagePos(state.doc));
  if (!node) throw new Error("the image is gone");
  return node;
}

describe("writing a new size back onto the image", () => {
  const opened = () => {
    const { doc } = importDocx(makeImageDocx(SEAL));
    return createEditorState(doc);
  };

  it("records the size in EMU and leaves everything else on the node alone", () => {
    const state = opened();
    const at = imagePos(state.doc);
    const before = imageAt(state);
    const tr = buildResizeImageTransaction(state, at, {
      width: 300,
      height: 150,
    });
    if (!tr) throw new Error("no transaction came back");

    const after = imageAt(state.apply(tr));
    expect(after.attrs.extent).toEqual({
      cx: pxToEmu(300),
      cy: pxToEmu(150),
    });
    expect(after.attrs.src).toBe(before.attrs.src);
    expect(after.attrs.alt).toBe(before.attrs.alt);
    // The original drawing XML stays, so only its two extents are rewritten on export
    expect(after.attrs.xml).toBe(before.attrs.xml);
  });

  it("keeps the run the picture sits in", () => {
    const state = opened();
    const tr = buildResizeImageTransaction(state, imagePos(state.doc), {
      width: 300,
      height: 150,
    });
    if (!tr) throw new Error("no transaction came back");

    const mark = imageAt(state.apply(tr)).marks.find(
      (entry) => entry.type.name === "run"
    );
    expect(mark?.attrs.rPr).toBe("<w:rPr><w:noProof/></w:rPr>");
  });

  it("makes no transaction for a drag that ended where it started", () => {
    const state = opened();
    expect(
      buildResizeImageTransaction(state, imagePos(state.doc), {
        width: 200,
        height: 100,
      })
    ).toBeNull();
  });

  it("makes no transaction where there is no image", () => {
    const state = opened();
    expect(
      buildResizeImageTransaction(state, 0, { width: 300, height: 150 })
    ).toBeNull();
  });
});

let mounted: (() => void)[] = [];

afterEach(() => {
  for (const dispose of mounted) dispose();
  mounted = [];
});

function openEditor(readOnly = false): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  const { doc, session } = importDocx(makeImageDocx(SEAL));
  const view = createEditorView({
    mount,
    state: createEditorState(doc),
    defaults: session.defaults,
    readOnly,
    onStateChange: () => undefined,
  });
  mounted.push(() => {
    view.destroy();
    mount.remove();
  });
  return view;
}

function selectImage(view: EditorView): void {
  view.dispatch(
    view.state.tr.setSelection(
      NodeSelection.create(view.state.doc, imagePos(view.state.doc))
    )
  );
}

function box(view: EditorView): HTMLElement {
  const found = view.dom.querySelector(`.${editorClassNames.imageBox}`);
  if (!(found instanceof HTMLElement)) throw new Error("no image box");
  return found;
}

function handle(view: EditorView, corner: ImageCorner): HTMLElement {
  const found = box(view).querySelector(`[data-corner="${corner}"]`);
  if (!(found instanceof HTMLElement)) throw new Error(`no ${corner} handle`);
  return found;
}

function drawnImage(view: EditorView): HTMLImageElement {
  const found = box(view).querySelector("img");
  if (!found) throw new Error("no image element");
  return found;
}

function press(element: HTMLElement, x: number, y: number): void {
  element.dispatchEvent(
    new MouseEvent("mousedown", {
      bubbles: true,
      cancelable: true,
      button: 0,
      clientX: x,
      clientY: y,
    })
  );
}

function move(x: number, y: number): void {
  window.dispatchEvent(new MouseEvent("mousemove", { clientX: x, clientY: y }));
}

function release(): void {
  window.dispatchEvent(new MouseEvent("mouseup"));
}

function pressEscape(): void {
  window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
}

describe("dragging a corner of a selected image", () => {
  it("draws the image inside a box that carries the four handles", () => {
    const view = openEditor();
    const corners = Array.from(box(view).children)
      .map((child) => child.getAttribute("data-corner"))
      .filter((corner) => corner !== null);

    expect(drawnImage(view).getAttribute("width")).toBe("200");
    expect(corners).toEqual(["nw", "ne", "sw", "se"]);
  });

  it("marks the box only while the image is the selection", () => {
    const view = openEditor();
    expect(box(view).classList.contains(editorClassNames.imageSelected)).toBe(
      false
    );

    selectImage(view);
    expect(box(view).classList.contains(editorClassNames.imageSelected)).toBe(
      true
    );

    // The caret moving into the text takes the mark off again
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1))
    );
    expect(box(view).classList.contains(editorClassNames.imageSelected)).toBe(
      false
    );
  });

  it("shows the new size while dragging and writes it down on release", () => {
    const view = openEditor();
    selectImage(view);

    press(handle(view, "se"), 300, 200);
    move(400, 250);
    // Only the drawn size moves while the button is down
    expect(drawnImage(view).getAttribute("width")).toBe("300");
    expect(imageAt(view.state).attrs.extent).toEqual({
      cx: 1905000,
      cy: 952500,
    });

    release();
    expect(imageAt(view.state).attrs.extent).toEqual({
      cx: pxToEmu(300),
      cy: pxToEmu(150),
    });
    expect(drawnImage(view).getAttribute("width")).toBe("300");
  });

  it("leaves a single step in the history however far the pointer travelled", () => {
    const view = openEditor();
    selectImage(view);

    press(handle(view, "se"), 300, 200);
    move(340, 220);
    move(370, 235);
    move(400, 250);
    release();
    expect(imageAt(view.state).attrs.extent.cx).toBe(pxToEmu(300));

    undo(view.state, (tr) => view.dispatch(tr));
    expect(imageAt(view.state).attrs.extent).toEqual({
      cx: 1905000,
      cy: 952500,
    });
  });

  it("leaves the image as it was when the drag is abandoned with Escape", () => {
    const view = openEditor();
    selectImage(view);

    press(handle(view, "se"), 300, 200);
    move(400, 250);
    pressEscape();
    // The preview is taken down and nothing is written
    expect(drawnImage(view).getAttribute("width")).toBe("200");
    expect(imageAt(view.state).attrs.extent).toEqual({
      cx: 1905000,
      cy: 952500,
    });

    // The pointer moving on afterwards no longer resizes anything
    move(600, 400);
    expect(drawnImage(view).getAttribute("width")).toBe("200");
  });

  it("neither shows handles nor resizes in a read-only editor", () => {
    const view = openEditor(true);
    selectImage(view);
    expect(box(view).classList.contains(editorClassNames.imageSelected)).toBe(
      false
    );

    press(handle(view, "se"), 300, 200);
    move(400, 250);
    release();
    expect(imageAt(view.state).attrs.extent).toEqual({
      cx: 1905000,
      cy: 952500,
    });
  });
});
