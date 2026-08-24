// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { LETTER_GEOMETRY, LETTER_SECT_PR, makeDocx } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { A4_BODY_WIDTH, A4_PORTRAIT } from "../docx/pageGeometry";
import { createEditorState } from "./createEditor";
import {
  documentBodyWidthPx,
  documentDefaultTabStopPt,
  documentGeometry,
} from "./documentStyles";

const BODY = '<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>';

/** A state built the way the editor builds one for this document */
function opened(sectPr: string) {
  const { doc, session } = importDocx(makeDocx(BODY + sectPr));
  return createEditorState(doc, { geometry: session.geometry });
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
    expect(documentDefaultTabStopPt(createEditorState(doc))).toBe(36);
    expect(
      documentDefaultTabStopPt(createEditorState(doc, { defaultTabStopPt: 48 }))
    ).toBe(48);
  });
});
