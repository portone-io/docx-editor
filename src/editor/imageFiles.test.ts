// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  LETTER_FIXTURE,
  LETTER_SECT_PR,
  makeDocx,
  readFixture,
  TINY_PNG,
} from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { A4_BODY_WIDTH } from "../docx/pageGeometry";
import type { SessionStore } from "../docx/session";
import { emuToPx, pxToEmu, toImageExtent } from "../ooxml/image";
import { createEditorState, createEditorView } from "./createEditor";
import {
  boundedImageSrc,
  fittedExtent,
  IMAGE_FILE_ACCEPT,
  imageFilesIn,
  insertImageFiles,
  readImageFile,
} from "./imageFiles";

/** A committed document on A4, to set beside the one on Letter */
const A4_FIXTURE = "kitchen-sink.docx";

function pngFile(name = "picture.png"): File {
  return new File([TINY_PNG], name, { type: "image/png" });
}

function textFile(): File {
  return new File(["notes"], "notes.txt", { type: "text/plain" });
}

describe("the file kinds we take", () => {
  it("asks a picker for exactly the kinds we can write back", () => {
    expect(IMAGE_FILE_ACCEPT).toBe("image/png,image/jpeg,image/gif,image/bmp");
  });

  it("picks out the image files and leaves the rest", () => {
    const png = pngFile();
    const gif = new File([], "moving.gif", { type: "image/gif" });
    const carrier = { files: [textFile(), png, gif] };

    expect(imageFilesIn(carrier)).toEqual([png, gif]);
  });

  it("finds nothing where there are no files at all", () => {
    expect(imageFilesIn({ files: [textFile()] })).toEqual([]);
    expect(imageFilesIn({ files: null })).toEqual([]);
    expect(imageFilesIn(null)).toEqual([]);
    expect(imageFilesIn(undefined)).toEqual([]);
  });

  it("reads nothing out of a file that is not an image kind we can carry", async () => {
    expect(await readImageFile(textFile())).toBeNull();
  });
});

describe("the size an image comes in at", () => {
  it("keeps the pixel size of an image that fits on the page", () => {
    expect(fittedExtent(200, 100)).toEqual({
      cx: pxToEmu(200),
      cy: pxToEmu(100),
    });
  });

  it("shrinks an image wider than the page's text width, ratio and all", () => {
    const extent = fittedExtent(4000, 3000);
    if (!extent) throw new Error("no size came back");

    expect(emuToPx(extent.cx)).toBeCloseTo(A4_BODY_WIDTH.px, 4);
    expect(extent.cx / extent.cy).toBeCloseTo(4 / 3, 4);
  });

  it("never grows an image narrower than the page", () => {
    const extent = fittedExtent(10, 10);
    expect(extent).toEqual({ cx: pxToEmu(10), cy: pxToEmu(10) });
  });

  it("takes the cap it is handed instead of the page width", () => {
    const extent = fittedExtent(400, 200, 100);
    expect(extent).toEqual({ cx: pxToEmu(100), cy: pxToEmu(50) });
  });

  it("fits a tall image within both dimensions when both caps are given", () => {
    expect(fittedExtent(100, 1000, 100, 200)).toEqual({
      cx: pxToEmu(20),
      cy: pxToEmu(200),
    });
  });

  it("gives back nothing for a size that cannot be shown", () => {
    expect(fittedExtent(0, 100)).toBeNull();
    expect(fittedExtent(100, 0)).toBeNull();
    expect(fittedExtent(Number.NaN, 100)).toBeNull();
    expect(fittedExtent(Number.POSITIVE_INFINITY, 100)).toBeNull();
    expect(fittedExtent(999_999_999_999_999, 1)).toBeNull();
  });
});

describe("embedded image bytes", () => {
  it("normalizes the declared MIME to the bytes' signature", () => {
    expect(boundedImageSrc("data:image/png;base64,/9j/2Q==")).toBe(
      "data:image/jpeg;base64,/9j/2Q=="
    );
  });

  it("rejects bytes with no supported image signature", () => {
    expect(boundedImageSrc("data:image/png;base64,AAAA")).toBeNull();
  });
});

let mounted: (() => void)[] = [];

afterEach(() => {
  for (const dispose of mounted) dispose();
  mounted = [];
  vi.unstubAllGlobals();
});

/**
 * Stands in for the browser decoding an image, which jsdom does not do: every image loaded
 * from here on reports this pixel size, so the width a file comes in at can be watched.
 */
function decodesAs(width: number, height: number): void {
  class DecodedImage {
    naturalWidth = width;
    naturalHeight = height;
    private loaded: (() => void)[] = [];

    addEventListener(name: string, handler: () => void): void {
      if (name === "load") this.loaded.push(handler);
    }

    removeEventListener(name: string, handler: () => void): void {
      if (name === "load") {
        this.loaded = this.loaded.filter((it) => it !== handler);
      }
    }

    set src(_value: string) {
      for (const handler of this.loaded) queueMicrotask(handler);
    }
  }
  vi.stubGlobal("Image", DecodedImage);
}

/** The first image node in the document */
function firstImage(doc: PMNode): PMNode {
  let found: PMNode | null = null;
  doc.descendants((node) => {
    if (!found && node.type.name === "image") found = node;
  });
  if (!found) throw new Error("no image went in");
  return found;
}

/** The width the drawing records, in pixels */
function widthPx(image: PMNode): number {
  const extent = toImageExtent(image.attrs.extent);
  if (!extent) throw new Error("the image records no size");
  return emuToPx(extent.cx);
}

/** An editor over an opened document, drawn on the paper that document names */
function mountEditor(opened: {
  doc: PMNode;
  session: SessionStore;
}): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  const view = createEditorView({
    mount,
    state: createEditorState(opened.doc, { geometry: opened.session.geometry }),
    defaults: opened.session.defaults,
    readOnly: false,
    onStateChange: () => undefined,
  });
  mounted.push(() => {
    view.destroy();
    mount.remove();
  });
  return view;
}

/**
 * An editor holding one paragraph, with the caret at the start of it.
 * `sectPr` is the section the document ends on, which is the paper it is drawn on
 */
function openEditor(sectPr = ""): EditorView {
  const view = mountEditor(
    importDocx(
      makeDocx(
        `<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>${sectPr}`
      )
    )
  );
  // jsdom does not measure text, so a position cannot be found from coordinates
  view.posAtCoords = () => ({ pos: 1, inside: 0 });
  return view;
}

/** An editor over one of the committed documents */
function openFixture(name: string): EditorView {
  return mountEditor(importDocx(readFixture(name)));
}

interface Transferred {
  text?: string;
  files?: File[];
}

/** Fakes one clipboard or drag payload. jsdom has neither DataTransfer nor ClipboardEvent */
function transfer({ text = "", files = [] }: Transferred) {
  return {
    types: ["text/plain"],
    files,
    getData: (type: string) => (type === "text/plain" ? text : ""),
  };
}

function fire(view: EditorView, name: string, property: string, value: object) {
  const event = new Event(name, { bubbles: true, cancelable: true });
  Object.defineProperty(event, property, { value });
  view.dom.dispatchEvent(event);
  return event.defaultPrevented;
}

function paste(view: EditorView, carried: Transferred): boolean {
  return fire(view, "paste", "clipboardData", transfer(carried));
}

function drop(view: EditorView, carried: Transferred): boolean {
  return fire(view, "drop", "dataTransfer", transfer(carried));
}

/**
 * Decoding an image is something jsdom cannot do, and these tests do not stand in for it
 * (`decodesAs` above), so a file that reaches the insert path never turns into a node here.
 * What can be told apart is whether the event was taken by the image path or fell through
 * to the plain-text one.
 */
describe("pasting and dropping image files", () => {
  it("takes a pasted image file, and the text of the same paste does not go in", () => {
    const view = openEditor();
    expect(paste(view, { text: "picture.png", files: [pngFile()] })).toBe(true);
    expect(view.state.doc.textContent).toBe("Body");
  });

  it("takes a dropped image file, and the text of the same drop does not go in", () => {
    const view = openEditor();
    expect(drop(view, { text: "picture.png", files: [pngFile()] })).toBe(true);
    expect(view.state.doc.textContent).toBe("Body");
  });

  it("leaves a paste that carries no image file to the plain-text path", () => {
    const view = openEditor();
    expect(paste(view, { text: "pasted ", files: [textFile()] })).toBe(true);
    expect(view.state.doc.textContent).toBe("pasted Body");
  });

  it("leaves a drop that carries no image file to the plain-text path", () => {
    const view = openEditor();
    expect(drop(view, { text: "dropped ", files: [textFile()] })).toBe(true);
    expect(view.state.doc.textContent).toBe("dropped Body");
  });

  it("puts nothing in for a file it cannot read", async () => {
    const view = openEditor();
    await insertImageFiles(view, [textFile()]);

    expect(view.state.doc.textContent).toBe("Body");
    expect(view.state.doc.child(0).childCount).toBe(1);
  });
});

describe("the width an inserted image is fitted to", () => {
  /** Wider than either paper's body, so the width it comes in at is the paper's own */
  const WIDE = 4000;

  it("is the body width of the paper the document names", async () => {
    decodesAs(WIDE, WIDE / 2);
    const view = openEditor(LETTER_SECT_PR);
    await insertImageFiles(view, [pngFile()]);

    // Letter with an inch of margin leaves 6.5in of body, which CSS draws as 624px
    const image = firstImage(view.state.doc);
    expect(widthPx(image)).toBeCloseTo(624, 4);
    expect(widthPx(image)).not.toBeCloseTo(A4_BODY_WIDTH.px, 1);
    // The ratio the file came in at is kept whichever paper it lands on
    const extent = toImageExtent(image.attrs.extent);
    expect(extent && extent.cx / extent.cy).toBeCloseTo(2, 4);
  });

  it("is A4's on a document that names no paper", async () => {
    decodesAs(WIDE, WIDE / 2);
    const view = openEditor();
    await insertImageFiles(view, [pngFile()]);

    expect(widthPx(firstImage(view.state.doc))).toBeCloseTo(
      A4_BODY_WIDTH.px,
      4
    );
  });

  it("leaves an image that fits the page at its own size", async () => {
    decodesAs(200, 100);
    const view = openEditor(LETTER_SECT_PR);
    await insertImageFiles(view, [pngFile()]);

    expect(widthPx(firstImage(view.state.doc))).toBe(200);
  });

  /**
   * The same file into two committed documents. Neither width is the A4 fallback's 627.47px:
   * each is the text column of the paper its own document is written on.
   */
  it("is a different width in a Letter document and in an A4 one", async () => {
    decodesAs(WIDE, WIDE / 2);

    const letter = openFixture(LETTER_FIXTURE);
    await insertImageFiles(letter, [pngFile()]);
    expect(widthPx(firstImage(letter.state.doc))).toBeCloseTo(624, 4);

    const a4 = openFixture(A4_FIXTURE);
    await insertImageFiles(a4, [pngFile()]);
    // 9026 twips of body at 96 pixels to the inch
    expect(widthPx(firstImage(a4.state.doc))).toBeCloseTo(601.7333, 3);
    expect(widthPx(firstImage(a4.state.doc))).not.toBeCloseTo(
      A4_BODY_WIDTH.px,
      1
    );
  });
});
