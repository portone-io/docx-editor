// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeStyledDocx,
} from "../../__testing__/docx";
import { posOfText, runCommand, select } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import type { LineSpacing } from "../../model/format";
import { createEditorState } from "../createEditor";
import {
  activeLineSpacing,
  canSetLineSpacing,
  SINGLE_LINE_SPACING,
  setLineSpacing,
} from "./spacingCommands";

const ONE_AND_A_HALF: LineSpacing = { rule: "auto", lines: 1.5 };
const DOUBLE: LineSpacing = { rule: "auto", lines: 2 };

const TEXT = "Body text";

function paragraph(text: string, pPr = ""): string {
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

interface Opened {
  state: EditorState;
  session: SessionStore;
}

function opened(body: string): Opened {
  const { doc, session } = importDocx(makeDocx(body));
  return { state: createEditorState(doc, { styles: session.styles }), session };
}

/** A document whose docDefaults declare a line spacing of their own */
function openedWithDefaults(body: string, pPrDefault: string): Opened {
  const { doc, session } = importDocx(
    makeStyledDocx(
      body,
      `<w:docDefaults><w:pPrDefault>${pPrDefault}</w:pPrDefault></w:docDefaults>`
    )
  );
  return {
    state: createEditorState(doc, {
      styles: session.styles,
      defaults: session.defaults,
    }),
    session,
  };
}

/** A state with the caret placed on the first character of the given text */
function at(state: EditorState, text: string): EditorState {
  return select(state, posOfText(state.doc, text));
}

function pPrOf(doc: PMNode, index: number): string {
  const pPr: unknown = doc.child(index).attrs.pPr;
  return typeof pPr === "string" ? pPr : "";
}

describe("setting the line spacing", () => {
  it("writes the preset onto the selected paragraph", () => {
    const { state, session } = opened(paragraph(TEXT));
    const spaced = runCommand(at(state, TEXT), setLineSpacing(ONE_AND_A_HALF));

    expect(documentXmlOf(spaced.doc, session)).toContain(
      '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
    );
  });

  it("keeps the space around the paragraph while the line spacing changes", () => {
    const pPr = '<w:pPr><w:spacing w:before="120" w:after="240"/></w:pPr>';
    const { state, session } = opened(paragraph(TEXT, pPr));
    const spaced = runCommand(at(state, TEXT), setLineSpacing(DOUBLE));

    expect(documentXmlOf(spaced.doc, session)).toContain(
      '<w:spacing w:before="120" w:after="240" w:line="480" w:lineRule="auto"/>'
    );
  });

  it("does nothing to a paragraph already drawn with that spacing", () => {
    const pPr = '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>';
    const { state } = opened(paragraph(TEXT, pPr));
    const spot = at(state, TEXT);

    expect(setLineSpacing(ONE_AND_A_HALF)(spot)).toBe(false);
    expect(setLineSpacing(DOUBLE)(spot)).toBe(true);
  });

  it("all the selected paragraphs change together", () => {
    const { state } = opened(paragraph("First") + paragraph("Second"));
    const both = select(
      state,
      posOfText(state.doc, "First"),
      posOfText(state.doc, "Second")
    );
    const spaced = runCommand(both, setLineSpacing(DOUBLE));

    expect(pPrOf(spaced.doc, 0)).toContain('w:line="480"');
    expect(pPrOf(spaced.doc, 1)).toContain('w:line="480"');
  });
});

describe("deciding the active line spacing", () => {
  it("reads a paragraph where nobody wrote a spacing as a single line", () => {
    const { state } = opened(paragraph(TEXT));
    expect(activeLineSpacing(at(state, TEXT))).toEqual(SINGLE_LINE_SPACING);
    expect(setLineSpacing(SINGLE_LINE_SPACING)(at(state, TEXT))).toBe(false);
  });

  it("reports the spacing the paragraph wrote down", () => {
    const pPr = '<w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>';
    const { state } = opened(paragraph(TEXT, pPr));
    expect(activeLineSpacing(at(state, TEXT))).toEqual(DOUBLE);
  });

  it("reports a pinned-down height as it is, matching no preset", () => {
    const pPr = '<w:pPr><w:spacing w:line="360" w:lineRule="exact"/></w:pPr>';
    const { state } = opened(paragraph(TEXT, pPr));
    expect(activeLineSpacing(at(state, TEXT))).toEqual({
      rule: "exact",
      pt: 18,
    });
  });

  it("reports nothing when the spacings are mixed", () => {
    const pPr = '<w:pPr><w:spacing w:line="480" w:lineRule="auto"/></w:pPr>';
    const { state } = opened(paragraph("First", pPr) + paragraph("Second"));
    const both = select(
      state,
      posOfText(state.doc, "First"),
      posOfText(state.doc, "Second")
    );
    expect(activeLineSpacing(both)).toBeNull();
  });

  it("the spacing a style gives counts toward the active decision", () => {
    const { doc, session } = importDocx(
      makeStyledDocx(
        paragraph(TEXT, '<w:pPr><w:pStyle w:val="Wide"/></w:pPr>'),
        '<w:style w:styleId="Wide"><w:pPr>' +
          '<w:spacing w:line="480" w:lineRule="auto"/></w:pPr></w:style>'
      )
    );
    const state = createEditorState(doc, { styles: session.styles });
    expect(activeLineSpacing(at(state, TEXT))).toEqual(DOUBLE);
  });

  it("the document default is what a paragraph with no spacing of its own falls back on", () => {
    const { state } = openedWithDefaults(
      paragraph(TEXT),
      '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
    );
    expect(activeLineSpacing(at(state, TEXT))).toEqual(ONE_AND_A_HALF);
  });

  it("there is a paragraph to space out wherever the caret sits in one", () => {
    const { state } = opened(paragraph(TEXT));
    expect(canSetLineSpacing(at(state, TEXT))).toBe(true);
  });

  it("there is none where the selection reaches no paragraph at all", () => {
    const { state } = opened("<w:customXml/>");
    expect(canSetLineSpacing(state)).toBe(false);
    expect(setLineSpacing(DOUBLE)(state)).toBe(false);
  });
});
