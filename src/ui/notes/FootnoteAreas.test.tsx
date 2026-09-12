// @vitest-environment jsdom
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { makeFormattedNotesDocx, makeNotesDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { noteProjection } from "../../editor/commands/noteQueries";
import {
  createEditorState,
  createEditorView,
  editorStateForSession,
} from "../../editor/createEditor";
import { documentOf, storyDocument } from "../../editor/editorDocument";
import {
  footnoteExtensions,
  footnoteHost,
} from "../../editor/notes/footnoteSurface";
import type { StoryCaret } from "../../editor/stories/storyView";
import { FOOTNOTE_BAND } from "../../page/demands/footnoteDemands";
import type {
  PageFace,
  PageOverlay,
  ReservedRoom,
} from "../../page/usePageLayout";
import { editorClassNames } from "../../styles/classNames";
import { DEFAULT_FONT_FALLBACKS } from "../../styles/fontStack";
import { FootnoteAreas, type FootnoteAreasProps } from "./FootnoteAreas";
import type { RowEditing } from "./StoryRow";

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

/** Footnote 2 carries a bold run and a second paragraph and is labelled 1; footnote 5 is labelled 2 */
function openedFootnotes() {
  const state = createEditorState(importDocx(makeFormattedNotesDocx()).doc);
  return noteProjection.read(state).footnotes;
}

/** A reference drawing a mark of its own, which the editor draws no number for */
const CUSTOM_MARK_BODY =
  '<w:p><w:r><w:t xml:space="preserve">Marked</w:t></w:r>' +
  '<w:r><w:footnoteReference w:customMarkFollows="1" w:id="2"/></w:r>' +
  "<w:r><w:t>*</w:t></w:r></w:p>";

function customMarkFootnotes() {
  const state = createEditorState(
    importDocx(makeNotesDocx(CUSTOM_MARK_BODY)).doc
  );
  return noteProjection.read(state).footnotes;
}

function room(
  ids: readonly string[],
  top: number,
  clipped = false
): ReservedRoom {
  return { band: FOOTNOTE_BAND, ids, top, height: 80, clipped };
}

function face(page: number, reserved: readonly ReservedRoom[]): PageFace {
  return {
    page,
    pos: 0,
    pageInSection: page,
    headerTop: 0,
    footerTop: 0,
    left: 80,
    width: 640,
    crossed: false,
    reserved,
  };
}

function overlayOf(...pages: PageFace[]): PageOverlay {
  return { left: 0, top: 0, width: 800, sheetHeight: 4000, marks: [], pages };
}

function draw(
  props: Pick<FootnoteAreasProps, "overlay"> & Partial<FootnoteAreasProps>
): void {
  const live = createRoot(host);
  root = live;
  act(() =>
    live.render(
      <FootnoteAreas
        footnotes={openedFootnotes()}
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
    `section[aria-label="Footnotes on page ${page}"]`
  );
  if (!(found instanceof HTMLElement)) {
    throw new Error(`no footnotes were drawn on page ${page}`);
  }
  return found;
}

function rowsOf(element: HTMLElement): HTMLElement[] {
  return [
    ...element.querySelectorAll<HTMLElement>(`.${editorClassNames.noteRow}`),
  ];
}

describe("the footnotes at the foot of each page", () => {
  it("draws each footnote under the page its reference stands on", () => {
    draw({
      overlay: overlayOf(
        face(1, [room(["footnote:2"], 900)]),
        face(2, []),
        face(3, [room(["footnote:5"], 3100)])
      ),
    });

    expect(area(1).textContent).toContain("bold words");
    expect(area(1).textContent).not.toContain("Later footnote");
    expect(area(3).textContent).toContain("Later footnote");
    expect(area(3).style).toMatchObject({
      top: "3100px",
      left: "80px",
      width: "640px",
      height: "80px",
    });
    expect(host.querySelectorAll("section")).toHaveLength(2);
  });

  it("draws a footnote's own reference mark as the reference's label", () => {
    draw({
      overlay: overlayOf(face(1, [room(["footnote:2", "footnote:5"], 900)])),
    });

    const marks = [
      ...area(1).querySelectorAll(`sup.${editorClassNames.noteMark}`),
    ].map((mark) => mark.textContent);
    expect(marks).toEqual(["1", "2"]);
    expect(area(1).textContent).not.toContain("footnoteRef");
  });

  it("names a footnote drawing a mark of its own by its kind alone", () => {
    draw({
      footnotes: customMarkFootnotes(),
      overlay: overlayOf(face(1, [room(["footnote:2"], 900)])),
    });

    // The reference draws no number, so there is none to name the row by
    expect(
      rowsOf(area(1)).map((row) => row.getAttribute("aria-label"))
    ).toEqual(["Footnote"]);
  });

  it("keeps a footnote's bold run and second paragraph", () => {
    draw({ overlay: overlayOf(face(1, [room(["footnote:2"], 900)])) });

    const [row] = rowsOf(area(1));
    if (!row) throw new Error("the footnote was not drawn");
    expect(
      [...row.querySelectorAll(`p.${editorClassNames.paragraph}`)].map(
        (paragraph) => paragraph.textContent
      )
    ).toEqual(["1Plain then bold words", "Second paragraph"]);
    const bold = [
      ...row.querySelectorAll<HTMLElement>(`.${editorClassNames.run}`),
    ].find((run) => run.textContent === "bold words");
    expect(bold?.style.fontWeight).toBe("bold");
  });

  it("names each footnote area for assistive technology", () => {
    draw({
      overlay: overlayOf(
        face(1, [room(["footnote:2"], 900)]),
        face(2, [room(["footnote:5"], 2000)])
      ),
    });

    expect(
      [...host.querySelectorAll("section")].map((section) =>
        section.getAttribute("aria-label")
      )
    ).toEqual(["Footnotes on page 1", "Footnotes on page 2"]);
  });

  it("draws no editor view for any note", () => {
    draw({
      overlay: overlayOf(
        face(1, [room(["footnote:2"], 900)]),
        face(2, [room(["footnote:5"], 2000)])
      ),
    });

    expect(rowsOf(host)).toHaveLength(2);
    expect(host.querySelector("[contenteditable]")).toBeNull();
    expect(host.querySelector(".ProseMirror")).toBeNull();
  });

  it("keeps a footnote out of sight until its height is known", () => {
    draw({
      overlay: overlayOf(face(1, [room(["footnote:2", "footnote:5"], 900)])),
      heights: new Map([["footnote:2", 40]]),
    });

    expect(rowsOf(area(1)).map((row) => row.style.visibility)).toEqual([
      "",
      "hidden",
    ]);
  });

  it("scrolls the footnotes of a page that holds less room than they ask for", () => {
    draw({
      overlay: overlayOf(
        face(1, [room(["footnote:2"], 900)]),
        face(2, [room(["footnote:5"], 1300, true)])
      ),
    });

    expect(area(1).style.overflowY).toBe("hidden");
    expect(area(2).style.overflowY).toBe("auto");
  });

  it("mounts one editor view over the entered footnote and none after leaving it", () => {
    const main = createEditorView({
      mount: document.createElement("div"),
      state: editorStateForSession(importDocx(makeFormattedNotesDocx())),
      onStateChange: () => {},
    });
    const held: { current: StoryCaret | null } = { current: null };
    const editing: RowEditing = {
      host: footnoteHost(main, () => {}),
      document: storyDocument(
        documentOf(main.state),
        documentOf(main.state).geometry
      ),
      extensions: footnoteExtensions(main, "2", () => "1"),
      caret: {
        take: () => held.current,
        keep: (caret) => {
          held.current = caret;
        },
      },
      onStateChange: () => {},
    };
    const overlay = overlayOf(
      face(1, [room(["footnote:2", "footnote:5"], 900)])
    );

    draw({ overlay, open: "footnote:2", editing });

    const views = host.querySelectorAll(".ProseMirror");
    expect(views).toHaveLength(1);
    expect(rowsOf(area(1))[0]?.contains(views[0] ?? null)).toBe(true);
    expect(rowsOf(area(1))[0]?.className).toContain(
      editorClassNames.noteRowOpen
    );
    // The note still says what it said, its own number drawn as the label it carries
    expect(rowsOf(area(1))[0]?.textContent).toContain("1Plain then bold words");

    draw({ overlay });

    expect(host.querySelector(".ProseMirror")).toBeNull();
    expect(host.querySelector("[contenteditable]")).toBeNull();
    expect(rowsOf(area(1))[0]?.textContent).toContain("1Plain then bold words");
    main.destroy();
  });

  it("holds none of the document's own source in a selection made inside a footnote area", () => {
    draw({ overlay: overlayOf(face(1, [room(["footnote:2"], 900)])) });

    // What a selection's contents are is what the browser writes when it copies one by itself
    const range = document.createRange();
    range.selectNodeContents(area(1));
    const selected = document.createElement("div");
    selected.append(range.cloneContents());

    expect(selected.textContent).toContain("bold words");
    expect(selected.innerHTML).not.toContain("<w:");
    expect(
      selected.querySelector(
        "[data-rpr], [data-rattrs], [data-fmt], [data-ppr]"
      )
    ).toBeNull();
  });
});
