// @vitest-environment jsdom
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { makeFormattedNotesDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { noteProjection } from "../../editor/commands/noteQueries";
import { createEditorState } from "../../editor/createEditor";
import type {
  PageFace,
  PageOverlay,
  TrailingRoom,
} from "../../page/usePageLayout";
import { editorClassNames } from "../../styles/classNames";
import { DEFAULT_FONT_FALLBACKS } from "../../styles/fontStack";
import { TrailingNotes, type TrailingNotesProps } from "./TrailingNotes";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

let host: HTMLDivElement;
let root: Root | null = null;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  if (root) act(() => root?.unmount());
  root = null;
  host.remove();
});

/** The notes of the formatted fixture: two footnotes and one endnote in italics */
function openedRows() {
  const state = createEditorState(importDocx(makeFormattedNotesDocx()).doc);
  return noteProjection.read(state).rows;
}

function face(page: number, trailing: TrailingRoom | null): PageFace {
  return {
    page,
    pos: 0,
    pageInSection: page,
    headerTop: 0,
    footerTop: 0,
    left: 80,
    width: 640,
    crossed: false,
    reserved: [],
    trailing,
  };
}

function overlayOf(...pages: PageFace[]): PageOverlay {
  return { left: 0, top: 0, width: 800, sheetHeight: 4000, marks: [], pages };
}

function draw(
  props: Pick<TrailingNotesProps, "overlay"> & Partial<TrailingNotesProps>
): void {
  const live = createRoot(host);
  root = live;
  act(() =>
    live.render(
      <TrailingNotes
        rows={openedRows()}
        heights={new Map()}
        onHeight={() => {}}
        fontFallbacks={DEFAULT_FONT_FALLBACKS}
        textStyle={{}}
        zoom={1}
        {...props}
      />
    )
  );
}

function area(page: number): HTMLElement {
  const found = host.querySelector(
    `section[aria-label="Endnotes on page ${page}"]`
  );
  if (!(found instanceof HTMLElement)) {
    throw new Error(`no endnotes were drawn on page ${page}`);
  }
  return found;
}

function rowTexts(within: HTMLElement): (string | null)[] {
  return [...within.querySelectorAll(`.${editorClassNames.noteRow}`)].map(
    (row) => row.textContent
  );
}

describe("the endnotes after the last paragraph", () => {
  it("draws endnotes after the last paragraph on pages of their own", () => {
    draw({
      overlay: overlayOf(
        face(1, null),
        face(2, { ids: ["endnote:3"], top: 1400, height: 90 })
      ),
    });

    const drawn = area(2);
    expect(host.querySelector('section[aria-label="Endnotes on page 1"]')).toBe(
      null
    );
    expect(drawn.style.top).toBe("1400px");
    expect(drawn.style.left).toBe("80px");
    expect(drawn.style.width).toBe("640px");
    expect(rowTexts(drawn)).toEqual(["1 Italic endnote"]);
    // The note's number is its label, drawn the way the reference in the text draws it
    expect(
      drawn.querySelector(`sup.${editorClassNames.noteMark}`)?.textContent
    ).toBe("1");
  });

  it("draws the rule once, above the page the endnotes begin on", () => {
    draw({
      overlay: overlayOf(
        face(1, { ids: ["endnote:3"], top: 900, height: 90 }),
        face(2, { ids: ["footnote:2"], top: 100, height: 40 })
      ),
    });

    const rule = `.${editorClassNames.noteSeparator}`;
    expect(area(1).querySelectorAll(rule)).toHaveLength(1);
    expect(area(2).querySelectorAll(rule)).toHaveLength(0);
  });

  it("keeps a row nobody has measured yet unseen, and shows it once it is measured", () => {
    const page = face(1, { ids: ["endnote:3"], top: 900, height: 90 });
    draw({ overlay: overlayOf(page) });

    const hidden = area(1).querySelector<HTMLElement>(
      `.${editorClassNames.noteRow}`
    );
    expect(hidden?.style.visibility).toBe("hidden");

    draw({
      overlay: overlayOf(page),
      heights: new Map([["endnote:3", 90]]),
    });

    const shown = area(1).querySelector<HTMLElement>(
      `.${editorClassNames.noteRow}`
    );
    expect(shown?.style.visibility).toBe("");
  });
});
