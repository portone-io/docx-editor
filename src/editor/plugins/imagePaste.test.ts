// @vitest-environment jsdom

import { undo } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import { AllSelection, TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it, vi } from "vitest";
import {
  documentXmlOf,
  makeImageDocx,
  TINY_PNG,
  TINY_PNG_DATA_URL,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { MAX_IMAGE_BYTES } from "../../docx/media";
import type { SessionStore } from "../../docx/session";
import { emuToPx, pxToEmu, toImageExtent } from "../../ooxml/image";
import { editorClassNames } from "../../styles/classNames";
import { PASTED_IMAGE_ATTRIBUTE } from "../clipboard/images";
import { createEditorState, createEditorView } from "../createEditor";

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

const TABLE_BODY =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${cellXml("Left")}${cellXml("Right")}</w:tr>` +
  "</w:tbl>";

function openEditor(
  body = '<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>'
): { view: EditorView; session: SessionStore } {
  const { doc, session } = importDocx(makeImageDocx(body));
  const view = createEditorView({
    mount: document.createElement("div"),
    state: createEditorState(doc),
    defaults: session.defaults,
    readOnly: false,
    onStateChange: () => undefined,
  });
  view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
  return { view, session };
}

function firstImage(doc: PMNode): PMNode | null {
  let found: PMNode | null = null;
  doc.descendants((node) => {
    if (node.type.name === "image") {
      found = node;
      return false;
    }
    return found === null;
  });
  return found;
}

function imageCount(doc: PMNode): number {
  let count = 0;
  doc.descendants((node) => {
    if (node.type.name === "image") count += 1;
  });
  return count;
}

function cellPositions(doc: PMNode): number[] {
  const positions: number[] = [];
  doc.descendants((node, position) => {
    if (node.type.name === "tableCell") positions.push(position);
    return true;
  });
  return positions;
}

function cellTexts(doc: PMNode): string[] {
  const texts: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "tableCell") texts.push(node.textContent);
    return true;
  });
  return texts;
}

function paste(view: EditorView, html: string, text = "image"): boolean {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: {
      files: [],
      getData: (type: string) =>
        type === "text/html" ? html : type === "text/plain" ? text : "",
    },
  });
  view.dom.dispatchEvent(event);
  return event.defaultPrevented;
}

function decodesAs(width: number, height: number): void {
  class DecodedImage {
    naturalWidth = width;
    naturalHeight = height;
    private loaded: (() => void)[] = [];

    addEventListener(name: string, handler: () => void): void {
      if (name === "load") this.loaded.push(handler);
    }

    removeEventListener(name: string, handler: () => void): void {
      if (name === "load")
        this.loaded = this.loaded.filter((it) => it !== handler);
    }

    set src(_value: string) {
      for (const handler of this.loaded) queueMicrotask(handler);
    }
  }
  vi.stubGlobal("Image", DecodedImage);
}

function decodeFails(): void {
  class BrokenImage {
    private failed: (() => void)[] = [];

    addEventListener(name: string, handler: () => void): void {
      if (name === "error") this.failed.push(handler);
    }

    removeEventListener(name: string, handler: () => void): void {
      if (name === "error") {
        this.failed = this.failed.filter((it) => it !== handler);
      }
    }

    set src(_value: string) {
      for (const handler of this.failed) queueMicrotask(handler);
    }
  }
  vi.stubGlobal("Image", BrokenImage);
}

function imageResponse(type = "image/png"): Response {
  const body = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(TINY_PNG);
      controller.close();
    },
  });
  return {
    ok: true,
    headers: new Headers({
      "content-length": `${TINY_PNG.length}`,
      "content-type": type,
    }),
    body,
  } as Response;
}

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("pasting images carried by HTML", () => {
  it.each([
    ["readable", TINY_PNG_DATA_URL],
    ["unreadable", "https://cdn.example/missing.png"],
  ])(
    "leaves a %s image over multiple cells to the regular clipboard path",
    (_kind, source) => {
      const fetch = vi.fn();
      vi.stubGlobal("fetch", fetch);
      const { view } = openEditor(TABLE_BODY);
      const cells = cellPositions(view.state.doc);
      view.dispatch(
        view.state.tr.setSelection(
          CellSelection.create(view.state.doc, cells[0] ?? 0, cells[1] ?? 0)
        )
      );

      expect(paste(view, `<p>replacement<img src="${source}"></p>`)).toBe(true);

      expect(cellTexts(view.state.doc)).toEqual(["", "replacement"]);
      expect(firstImage(view.state.doc)).toBeNull();
      expect(fetch).not.toHaveBeenCalled();
      view.destroy();
    }
  );

  it("downloads a web image and inserts it with its natural size", async () => {
    decodesAs(320, 160);
    const fetch = vi.fn(async (_url: URL, _init: RequestInit) =>
      imageResponse()
    );
    vi.stubGlobal("fetch", fetch);
    const { view, session } = openEditor();

    expect(
      paste(
        view,
        '<p>Before <img src="https://cdn.example/seal.png" alt="seal"> after</p>'
      )
    ).toBe(true);
    expect(view.state.doc.textContent).toBe("Body");

    await vi.waitFor(() => expect(firstImage(view.state.doc)).not.toBeNull());
    const image = firstImage(view.state.doc);
    expect(view.state.doc.textContent).toBe("Before  after");
    expect(image?.attrs).toMatchObject({
      extent: { cx: pxToEmu(320), cy: pxToEmu(160) },
      alt: "seal",
      xml: null,
    });
    expect(`${fetch.mock.calls[0]?.[0]}`).toBe("https://cdn.example/seal.png");
    expect(fetch.mock.calls[0]?.[1]).toEqual(
      expect.objectContaining({
        credentials: "omit",
        referrerPolicy: "no-referrer",
        signal: expect.any(AbortSignal),
      })
    );
    const xml = documentXmlOf(view.state.doc, session);
    expect(xml).toContain(
      `<wp:extent cx="${pxToEmu(320)}" cy="${pxToEmu(160)}"/>`
    );
    expect(xml).toContain('descr="seal"');
    view.destroy();
  });

  it("uses the plain-text fallback when the web image cannot be read", async () => {
    const fetch = vi.fn(async () => imageResponse("text/plain"));
    vi.stubGlobal("fetch", fetch);
    const { view } = openEditor();

    expect(
      paste(view, '<p><img src="https://cdn.example/missing.png"></p>')
    ).toBe(true);
    await vi.waitFor(() => expect(fetch).toHaveBeenCalledOnce());
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(view.state.doc.textContent).toBe("image");
    expect(firstImage(view.state.doc)).toBeNull();
    view.destroy();
  });

  it("keeps the paste position mapped while the image is loading", async () => {
    decodesAs(40, 20);
    let resolveResponse: (response: Response) => void = () => undefined;
    const response = new Promise<Response>((resolve) => {
      resolveResponse = resolve;
    });
    vi.stubGlobal(
      "fetch",
      vi.fn(() => response)
    );
    const { view } = openEditor();
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1))
    );

    expect(paste(view, '<img src="https://cdn.example/delayed.png">')).toBe(
      true
    );
    view.dispatch(view.state.tr.insertText("typed"));
    resolveResponse(imageResponse());

    await vi.waitFor(() => expect(firstImage(view.state.doc)).not.toBeNull());
    const paragraph = view.state.doc.firstChild;
    expect(paragraph?.child(0).type.name).toBe("image");
    expect(paragraph?.textContent).toBe("typedBody");
    view.destroy();
  });

  it("cancels a selected-range paste when typing replaces that range", async () => {
    let signal: AbortSignal | null = null;
    vi.stubGlobal(
      "fetch",
      vi.fn(
        (_url: URL, init: RequestInit) =>
          new Promise<Response>((_resolve, reject) => {
            signal = init.signal as AbortSignal;
            signal.addEventListener("abort", () =>
              reject(new DOMException("aborted", "AbortError"))
            );
          })
      )
    );
    const { view } = openEditor();

    paste(view, '<img src="https://cdn.example/delayed.png">');
    view.dispatch(view.state.tr.insertText("typed"));

    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    expect(view.state.doc.textContent).toBe("typed");
    expect(firstImage(view.state.doc)).toBeNull();
    view.destroy();
  });

  it("cancels a selected-range paste when typing inside that range", async () => {
    let signal: AbortSignal | null = null;
    vi.stubGlobal(
      "fetch",
      vi.fn(
        (_url: URL, init: RequestInit) =>
          new Promise<Response>((_resolve, reject) => {
            signal = init.signal as AbortSignal;
            signal.addEventListener("abort", () =>
              reject(new DOMException("aborted", "AbortError"))
            );
          })
      )
    );
    const { view } = openEditor();
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 4))
    );

    paste(view, '<img src="https://cdn.example/delayed.png">');
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 2))
    );
    view.dispatch(view.state.tr.insertText("x"));

    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    expect(view.state.doc.textContent).toBe("Bxody");
    expect(firstImage(view.state.doc)).toBeNull();
    view.destroy();
  });

  it("cancels a selected-range paste when typing at its start", async () => {
    let signal: AbortSignal | null = null;
    vi.stubGlobal(
      "fetch",
      vi.fn(
        (_url: URL, init: RequestInit) =>
          new Promise<Response>((_resolve, reject) => {
            signal = init.signal as AbortSignal;
            signal.addEventListener("abort", () =>
              reject(new DOMException("aborted", "AbortError"))
            );
          })
      )
    );
    const { view } = openEditor();
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 4))
    );

    paste(view, '<img src="https://cdn.example/delayed.png">');
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1))
    );
    view.dispatch(view.state.tr.insertText("x"));

    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    expect(view.state.doc.textContent).toBe("xBody");
    expect(firstImage(view.state.doc)).toBeNull();
    view.destroy();
  });

  it("cancels a selected-range paste when formatting that range", async () => {
    let signal: AbortSignal | null = null;
    vi.stubGlobal(
      "fetch",
      vi.fn(
        (_url: URL, init: RequestInit) =>
          new Promise<Response>((_resolve, reject) => {
            signal = init.signal as AbortSignal;
            signal.addEventListener("abort", () =>
              reject(new DOMException("aborted", "AbortError"))
            );
          })
      )
    );
    const { view } = openEditor();
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 4))
    );
    const run = view.state.schema.marks.run?.create({
      rPr: "<w:rPr><w:b/></w:rPr>",
    });
    if (!run) throw new Error("the run mark is missing");

    paste(view, '<img src="https://cdn.example/delayed.png">');
    view.dispatch(view.state.tr.addMark(1, 4, run));

    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    expect(view.state.doc.textContent).toBe("Body");
    expect(view.state.doc.rangeHasMark(1, 4, run.type)).toBe(true);
    expect(firstImage(view.state.doc)).toBeNull();
    view.destroy();
  });

  it("cancels a pending paste when history moves", async () => {
    let signal: AbortSignal | null = null;
    vi.stubGlobal(
      "fetch",
      vi.fn(
        (_url: URL, init: RequestInit) =>
          new Promise<Response>((_resolve, reject) => {
            signal = init.signal as AbortSignal;
            signal.addEventListener("abort", () =>
              reject(new DOMException("aborted", "AbortError"))
            );
          })
      )
    );
    const { view } = openEditor();
    view.dispatch(
      view.state.tr
        .setSelection(TextSelection.create(view.state.doc, 1))
        .insertText("edited")
    );
    paste(view, '<img src="https://cdn.example/delayed.png">');

    expect(
      undo(view.state, (transaction) => view.dispatch(transaction), view)
    ).toBe(true);

    await vi.waitFor(() => expect(signal?.aborted).toBe(true));
    expect(firstImage(view.state.doc)).toBeNull();
    view.destroy();
  });

  it("aborts a pending download when the editor is destroyed", async () => {
    let signal: AbortSignal | null = null;
    vi.stubGlobal(
      "fetch",
      vi.fn(
        (_url: URL, init: RequestInit) =>
          new Promise<Response>((_resolve, reject) => {
            signal = init.signal as AbortSignal;
            signal.addEventListener("abort", () =>
              reject(new DOMException("aborted", "AbortError"))
            );
          })
      )
    );
    const { view } = openEditor();

    paste(view, '<img src="https://cdn.example/delayed.png">');
    await vi.waitFor(() => expect(signal).not.toBeNull());
    view.destroy();

    expect((signal as AbortSignal | null)?.aborted).toBe(true);
  });

  it("aborts a stalled download and inserts the fallback", async () => {
    vi.useFakeTimers();
    let signal: AbortSignal | null = null;
    vi.stubGlobal(
      "fetch",
      vi.fn(
        (_url: URL, init: RequestInit) =>
          new Promise<Response>((_resolve, reject) => {
            signal = init.signal as AbortSignal;
            signal.addEventListener("abort", () =>
              reject(new DOMException("aborted", "AbortError"))
            );
          })
      )
    );
    const { view } = openEditor();

    paste(view, '<img src="https://cdn.example/stalled.png">', "fallback");
    await vi.advanceTimersByTimeAsync(15_000);
    await Promise.resolve();

    expect((signal as AbortSignal | null)?.aborted).toBe(true);
    expect(view.state.doc.textContent).toBe("fallback");
    view.destroy();
  });

  it("times out image decoding as well as network loading", async () => {
    vi.useFakeTimers();
    let decodedSource = "not started";
    class StalledImage {
      addEventListener(): void {}
      removeEventListener(): void {}
      set src(value: string) {
        decodedSource = value;
      }
    }
    vi.stubGlobal("Image", StalledImage);
    const { view } = openEditor();

    paste(view, `<img src="${TINY_PNG_DATA_URL}">`, "fallback");
    await vi.advanceTimersByTimeAsync(15_000);
    await Promise.resolve();

    expect(view.state.doc.textContent).toBe("fallback");
    expect(firstImage(view.state.doc)).toBeNull();
    expect(decodedSource).toBe("");
    view.destroy();
  });

  it("lets the latest paste replace a pending paste at the same range", async () => {
    decodesAs(20, 10);
    const fetch = vi.fn((_url: URL, init: RequestInit) =>
      fetch.mock.calls.length === 1
        ? new Promise<Response>((_resolve, reject) => {
            const signal = init.signal as AbortSignal;
            signal.addEventListener("abort", () =>
              reject(new DOMException("aborted", "AbortError"))
            );
          })
        : Promise.resolve(imageResponse())
    );
    vi.stubGlobal("fetch", fetch);
    const { view } = openEditor();

    paste(view, '<img src="https://cdn.example/first.png" alt="first">');
    paste(view, '<img src="https://cdn.example/second.png" alt="second">');

    await vi.waitFor(() => expect(firstImage(view.state.doc)).not.toBeNull());
    expect(firstImage(view.state.doc)?.attrs.alt).toBe("second");
    expect(fetch).toHaveBeenCalledTimes(2);
    view.destroy();
  });

  it("lets an overlapping paste replace a pending paste", async () => {
    decodesAs(20, 10);
    let firstSignal: AbortSignal | null = null;
    const fetch = vi.fn((_url: URL, init: RequestInit) =>
      fetch.mock.calls.length === 1
        ? new Promise<Response>((_resolve, reject) => {
            firstSignal = init.signal as AbortSignal;
            firstSignal.addEventListener("abort", () =>
              reject(new DOMException("aborted", "AbortError"))
            );
          })
        : Promise.resolve(imageResponse())
    );
    vi.stubGlobal("fetch", fetch);
    const { view } = openEditor();
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 4))
    );

    paste(view, '<img src="https://cdn.example/first.png" alt="first">');
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 2, 5))
    );
    paste(view, '<img src="https://cdn.example/second.png" alt="second">');

    await vi.waitFor(() => expect(firstImage(view.state.doc)).not.toBeNull());
    expect((firstSignal as AbortSignal | null)?.aborted).toBe(true);
    expect(firstImage(view.state.doc)?.attrs.alt).toBe("second");
    expect(view.state.doc.textContent).toBe("B");
    view.destroy();
  });

  it("streams an undeclared response only up to the image limit", async () => {
    let canceled = false;
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(MAX_IMAGE_BYTES + 1));
      },
      cancel() {
        canceled = true;
      },
    });
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          ({
            ok: true,
            headers: new Headers({ "content-type": "image/png" }),
            body,
          }) as Response
      )
    );
    const { view } = openEditor();

    paste(view, '<img src="https://cdn.example/oversized.png">', "fallback");

    await vi.waitFor(() => expect(view.state.doc.textContent).toBe("fallback"));
    expect(canceled).toBe(true);
    expect(firstImage(view.state.doc)).toBeNull();
    view.destroy();
  });

  it("decodes copied editor bytes and bounds their requested extent", async () => {
    decodesAs(1, 1);
    const { view, session } = openEditor();

    paste(
      view,
      `<img ${PASTED_IMAGE_ATTRIBUTE}="0">` +
        `<img class="${editorClassNames.image}" src="${TINY_PNG_DATA_URL}" ` +
        'data-extent=\'{"cx":999999999999,"cy":999999999999}\' ' +
        'data-xml="source drawing">'
    );

    await vi.waitFor(() => expect(firstImage(view.state.doc)).not.toBeNull());
    const extent = toImageExtent(firstImage(view.state.doc)?.attrs.extent);
    expect(extent && emuToPx(extent.cx)).toBeLessThan(700);
    expect(extent && emuToPx(extent.cy)).toBeLessThan(1000);
    expect(firstImage(view.state.doc)?.attrs.xml).toBeNull();
    expect(imageCount(view.state.doc)).toBe(1);
    expect(documentXmlOf(view.state.doc, session)).not.toContain(
      "source drawing"
    );
    view.destroy();
  });

  it("rejects a copied extent that would round to an empty drawing", async () => {
    decodesAs(1, 1);
    const { view } = openEditor();

    paste(
      view,
      `<img class="${editorClassNames.image}" src="${TINY_PNG_DATA_URL}" ` +
        'data-extent=\'{"cx":999999999999999,"cy":1}\'>',
      "fallback"
    );

    await vi.waitFor(() => expect(view.state.doc.textContent).toBe("fallback"));
    expect(firstImage(view.state.doc)).toBeNull();
    view.destroy();
  });

  it("falls back when copied editor bytes do not decode", async () => {
    decodeFails();
    const { view } = openEditor();

    paste(
      view,
      `<img class="${editorClassNames.image}" src="${TINY_PNG_DATA_URL}" ` +
        'data-extent=\'{"cx":100,"cy":100}\'>',
      "fallback"
    );

    await vi.waitFor(() => expect(view.state.doc.textContent).toBe("fallback"));
    expect(firstImage(view.state.doc)).toBeNull();
    view.destroy();
  });

  it("does not fetch an unsafe image URL and keeps surrounding text", () => {
    const fetch = vi.fn();
    vi.stubGlobal("fetch", fetch);
    const { view } = openEditor();

    expect(
      paste(
        view,
        '<p>Before <img src="javascript:alert(1)"> after</p>',
        "Before after"
      )
    ).toBe(true);

    expect(view.state.doc.textContent).toBe("Before  after");
    expect(firstImage(view.state.doc)).toBeNull();
    expect(fetch).not.toHaveBeenCalled();
    view.destroy();
  });
});
