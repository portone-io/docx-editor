// @vitest-environment jsdom
import type { ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { makeDocx } from "./__testing__/docx";
import { renderInto } from "./__testing__/react";
import { DocxEditor, type DocxEditorHandle } from "./DocxEditor";
import {
  createDocxBlob,
  DOCX_MIME_TYPE,
  downloadBlob,
  downloadDocx,
  hasExportableContent,
  withDocxExtension,
} from "./download";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

const PARAGRAPH = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>'
);

/** A document with no text and no table, just one empty paragraph. The paper settings at the end are not an editable block */
const EMPTY = makeDocx(
  '<w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
);

const cellXml = (text: string) =>
  `<w:tc><w:p>${text === "" ? "" : `<w:r><w:t>${text}</w:t></w:r>`}</w:p></w:tc>`;

/** A table document whose cells are all empty, so it holds not a single character */
const EMPTY_TABLE = makeDocx(
  "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr>${cellXml("")}${cellXml("")}</w:tr>` +
    "</w:tbl>"
);

/** The anchor that was clicked, and whether it was attached to the document at the moment of the click */
interface Click {
  anchor: HTMLAnchorElement;
  connected: boolean;
}

let host: HTMLDivElement;
let created: string[];
let revoked: string[];
let clicks: Click[];

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);

  created = [];
  revoked = [];
  clicks = [];
  // The clicked anchor is taken off again immediately, so capture it at the moment of the click
  vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(
    function captured(this: HTMLAnchorElement) {
      clicks.push({ anchor: this, connected: this.isConnected });
    }
  );
  vi.stubGlobal("URL", {
    ...URL,
    createObjectURL: (blob: Blob) => {
      const url = `blob:test/${created.length}#${blob.size}`;
      created.push(url);
      return url;
    },
    revokeObjectURL: (url: string) => revoked.push(url),
  });
  vi.useFakeTimers();
});

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
  host.remove();
});

const render = (element: ReactNode) => renderInto(host, element);

function mount(bytes: Uint8Array) {
  const box: { current: DocxEditorHandle | null } = { current: null };
  const unmount = render(
    <DocxEditor document={bytes} ref={box} renderImportError={() => null} />
  );
  const handle = box.current;
  if (!handle) throw new Error("the ref was not attached");
  return { handle, unmount };
}

/** The download anchor that was just clicked */
function lastClick(): Click {
  const found = clicks[clicks.length - 1];
  if (!found) throw new Error("no anchor was clicked");
  return found;
}

describe("withDocxExtension", () => {
  // The name stays Korean here so that a non-ASCII file name is covered
  it("appends the extension when there is none", () => {
    expect(withDocxExtension("표준계약서")).toBe("표준계약서.docx");
  });

  it("does not append it twice when it is already there", () => {
    expect(withDocxExtension("standard-contract.docx")).toBe(
      "standard-contract.docx"
    );
  });

  it("normalizes an upper case extension too", () => {
    expect(withDocxExtension("standard-contract.DOCX")).toBe(
      "standard-contract.docx"
    );
  });

  it("leaves a docx in the middle of the name alone", () => {
    expect(withDocxExtension("docx.sample")).toBe("docx.sample.docx");
  });
});

describe("createDocxBlob", () => {
  it("carries the docx MIME type and the byte count as they are", () => {
    const blob = createDocxBlob(new Uint8Array([1, 2, 3]));
    expect(blob.type).toBe(DOCX_MIME_TYPE);
    expect(blob.size).toBe(3);
  });

  it("carries only the slice itself when given a slice of a larger buffer", () => {
    const whole = new Uint8Array([1, 2, 3, 4, 5]);
    expect(createDocxBlob(whole.subarray(1, 3)).size).toBe(2);
  });
});

describe("downloadBlob", () => {
  it("clicks an anchor to download under exactly that file name", () => {
    downloadBlob(new Blob(["x"]), "contract.docx");
    const { anchor } = lastClick();
    expect(anchor.download).toBe("contract.docx");
    expect(anchor.href).toBe(created[0]);
  });

  it("clicks the anchor while it is attached to the document and takes it off right away", () => {
    downloadBlob(new Blob(["x"]), "contract.docx");
    expect(lastClick().connected).toBe(true);
    expect(lastClick().anchor.isConnected).toBe(false);
  });

  it("holds the objectURL until the download has had time to start", () => {
    downloadBlob(new Blob(["x"]), "contract.docx");
    expect(revoked).toEqual([]);

    vi.advanceTimersByTime(60_000);
    expect(revoked).toEqual(created);
  });
});

describe("hasExportableContent", () => {
  it("has content when there is text", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    expect(hasExportableContent(handle.view.state.doc)).toBe(true);
    unmount();
  });

  it("is empty when there are only empty paragraphs", () => {
    const { handle, unmount } = mount(EMPTY);
    expect(hasExportableContent(handle.view.state.doc)).toBe(false);
    unmount();
  });

  it("has content when there is a table, even with no text", () => {
    const { handle, unmount } = mount(EMPTY_TABLE);
    expect(hasExportableContent(handle.view.state.doc)).toBe(true);
    unmount();
  });
});

describe("downloadDocx", () => {
  it("is unavailable when there is no editor", () => {
    expect(downloadDocx(null, { fileName: "contract" })).toEqual({
      status: "unavailable",
    });
    expect(created).toEqual([]);
  });

  it("is empty and makes no file when there is nothing to export", () => {
    const { handle, unmount } = mount(EMPTY);
    expect(downloadDocx(handle, { fileName: "contract" })).toEqual({
      status: "empty",
    });
    expect(created).toEqual([]);
    unmount();
  });

  it("returns the exported file name and byte count", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    const result = downloadDocx(handle, { fileName: "standard-contract" });
    expect(result).toEqual({
      status: "exported",
      fileName: "standard-contract.docx",
      byteLength: handle.exportBytes().length,
    });
    expect(lastClick().anchor.download).toBe("standard-contract.docx");
    unmount();
  });
});
