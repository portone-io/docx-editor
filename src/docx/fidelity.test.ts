// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { fidelityNotesOf } from "./fidelity";
import { importDocx } from "./importDocx";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

function notesOf(body: string) {
  const { doc, session } = importDocx(makeDocx(body));
  return fidelityNotesOf(doc, session.mainPartPath);
}

describe("the notes a document opens with", () => {
  it("reports a demoted paragraph as a placeholder note carrying its block number", () => {
    expect(
      notesOf(
        `<w:p>${run("First")}</w:p><w:p><w:r><w:sym w:char="F0E0"/></w:r></w:p>`
      )
    ).toEqual([
      {
        severity: "placeholder",
        code: "paragraph-demoted",
        part: "word/document.xml",
        block: 1,
        pos: 7,
        element: "w:p",
      },
    ]);
  });

  it("reports a demoted table as a placeholder note of its own", () => {
    // A row-level bookmark is a row child this reader has no node for, so the table is demoted
    const notes = notesOf(
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
        '<w:tr><w:bookmarkStart w:id="1" w:name="b" w:colFirst="0" w:colLast="0"/>' +
        `<w:tc><w:p>${run("a")}</w:p></w:tc><w:bookmarkEnd w:id="1"/></w:tr></w:tbl>`
    );
    expect(notes).toEqual([
      expect.objectContaining({
        severity: "placeholder",
        code: "table-demoted",
        block: 0,
        element: "w:tbl",
      }),
    ]);
  });

  it("reports a body-level bookmark as a hidden range marker", () => {
    expect(
      notesOf(
        '<w:bookmarkStart w:id="8" w:name="Appendix"/>' +
          `<w:p>${run("First")}</w:p>` +
          '<w:bookmarkEnd w:id="8"/>'
      )
    ).toEqual([
      expect.objectContaining({
        severity: "hidden",
        code: "range-marker",
        block: 0,
        element: "w:bookmarkStart",
      }),
      expect.objectContaining({
        severity: "hidden",
        code: "range-marker",
        block: 2,
        element: "w:bookmarkEnd",
      }),
    ]);
  });

  it("reports a bookmark inside a paragraph under the block it stands in", () => {
    expect(
      notesOf(
        `<w:p>${run("a")}<w:bookmarkStart w:id="1" w:name="Here"/>` +
          '<w:bookmarkEnd w:id="1"/></w:p>'
      )
    ).toEqual([
      expect.objectContaining({ code: "range-marker", block: 0, pos: 2 }),
      expect.objectContaining({ code: "range-marker", block: 0, pos: 3 }),
    ]);
  });

  it("reports an invisible paragraph child that is not a bookmark as hidden preserved inline", () => {
    expect(
      notesOf(
        `<w:p>${run("a")}<w:ins w:id="1" w:author="x" w:date="2026-01-01T00:00:00Z">` +
          `${run("inserted")}</w:ins>${run("b")}</w:p>`
      )
    ).toEqual([
      expect.objectContaining({
        severity: "hidden",
        code: "preserved-inline",
        block: 0,
        element: "w:ins",
      }),
    ]);
  });

  it("reports nothing for an empty run kept whole", () => {
    expect(
      notesOf(`<w:p>${run("a")}<w:r><w:rPr><w:b/></w:rPr></w:r></w:p>`)
    ).toEqual([]);
  });

  it("reports nothing for a document made of modelled nodes only", () => {
    expect(
      notesOf(
        `<w:p>${run("a")}</w:p>` +
          '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
          `<w:tr><w:tc><w:p>${run("b")}</w:p></w:tc></w:tr></w:tbl>`
      )
    ).toEqual([]);
  });

  it("leaves the part out when the caller holds no session", () => {
    const { doc } = importDocx(
      makeDocx(
        `<w:p>${run("a")}<w:bookmarkStart w:id="1" w:name="Here"/></w:p>`
      )
    );
    expect(fidelityNotesOf(doc, null)).toEqual([
      expect.objectContaining({ part: null, code: "range-marker" }),
    ]);
  });
});
