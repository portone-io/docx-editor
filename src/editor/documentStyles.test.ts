// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { LETTER_GEOMETRY, LETTER_SECT_PR, makeDocx } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { A4_BODY_WIDTH, A4_PORTRAIT } from "../docx/pageGeometry";
import { editorCssVariables } from "../styles/classNames";
import {
  createEditorState,
  createEditorView,
  editorStateForSession,
} from "./createEditor";
import {
  documentBodyWidthPx,
  documentDefaultTabStopPt,
  documentGeometry,
} from "./documentStyles";
import { type EditorDocument, NO_DOCUMENT } from "./editorDocument";

const BODY = '<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>';

/** A state built the way the editor builds one for this document */
function opened(sectPr: string) {
  const { doc, session } = importDocx(makeDocx(BODY + sectPr));
  return editorStateForSession({ doc, session });
}

describe("the paper the state carries", () => {
  it("is the one the document names", () => {
    expect(documentGeometry(opened(LETTER_SECT_PR))).toEqual(LETTER_GEOMETRY);
  });

  it("is A4 for a document that names none", () => {
    expect(documentGeometry(opened(""))).toEqual(A4_PORTRAIT);
  });

  it("is A4 for a state built without a document to ask", () => {
    const { doc } = importDocx(makeDocx(BODY + LETTER_SECT_PR));
    // The geometry is handed to the state, not read off the doc, so this one is the fallback
    expect(documentGeometry(createEditorState(doc))).toEqual(A4_PORTRAIT);
  });

  it("gives out the body width in the pixels an image is fitted to", () => {
    // Letter with an inch of margin leaves 6.5in of body, which CSS draws as 624px
    expect(documentBodyWidthPx(opened(LETTER_SECT_PR))).toBeCloseTo(624, 6);
    expect(documentBodyWidthPx(opened(""))).toBeCloseTo(A4_BODY_WIDTH.px, 6);
  });

  it("retains the document's automatic tab interval", () => {
    const { doc } = importDocx(makeDocx(BODY));
    const widerTabs: EditorDocument = { ...NO_DOCUMENT, defaultTabStopPt: 48 };
    expect(documentDefaultTabStopPt(createEditorState(doc))).toBe(36);
    expect(
      documentDefaultTabStopPt(createEditorState(doc, { document: widerTabs }))
    ).toBe(48);
  });
});

describe("the sheet drawn from the document snapshot", () => {
  it("updates the sheet from the state passed to the existing view", () => {
    const { doc } = importDocx(makeDocx(BODY));
    const replacement: EditorDocument = {
      ...NO_DOCUMENT,
      geometry: LETTER_GEOMETRY,
      defaults: {
        fontFamily: "Example Serif",
        fontSizePt: 18,
        lineSpacing: null,
      },
      defaultTabStopPt: 48,
    };
    const initial = createEditorState(doc);
    const next = createEditorState(doc, { document: replacement });
    const view = createEditorView({
      mount: document.createElement("div"),
      state: initial,
      onStateChange: () => {},
    });
    try {
      const originalStyle = view.dom.getAttribute("style");
      view.updateState(next);
      expect(
        view.dom.style.getPropertyValue(editorCssVariables.pageWidth)
      ).toBe("816px");
      expect(
        view.dom.style.getPropertyValue(editorCssVariables.pageHeight)
      ).toBe("1056px");
      expect(view.dom.style.getPropertyValue(editorCssVariables.fontSize)).toBe(
        "18pt"
      );
      expect(
        view.dom.style.getPropertyValue(editorCssVariables.fontFamily)
      ).toContain("Example Serif");
      expect(view.dom.style.tabSize).toBe("48pt");
      view.updateState(initial);
      expect(view.dom.getAttribute("style")).toBe(originalStyle);
    } finally {
      view.destroy();
    }
  });

  it("keeps font fallbacks local to views sharing one document snapshot", () => {
    const { doc } = importDocx(makeDocx(BODY));
    const state = createEditorState(doc);
    const views = ["serif", "monospace"].map((defaultStack) =>
      createEditorView({
        mount: document.createElement("div"),
        state,
        fontFallbacks: {
          groups: [],
          defaultStack,
          defaultFontName: defaultStack,
        },
        onStateChange: () => {},
      })
    );
    try {
      expect(
        views[0].dom.style.getPropertyValue(editorCssVariables.fontFamily)
      ).toBe("serif");
      expect(
        views[1].dom.style.getPropertyValue(editorCssVariables.fontFamily)
      ).toBe("monospace");
      views[0].updateState(state.apply(state.tr.insertText("A", 1)));
      expect(
        views[0].dom.style.getPropertyValue(editorCssVariables.fontFamily)
      ).toBe("serif");
      views[1].updateState(state.apply(state.tr.insertText("B", 1)));
      expect(
        views[1].dom.style.getPropertyValue(editorCssVariables.fontFamily)
      ).toBe("monospace");
    } finally {
      for (const view of views) view.destroy();
    }
  });
});
