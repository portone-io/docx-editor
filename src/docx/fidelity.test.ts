// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import {
  fixtureNames,
  makeDocx,
  producerFixtureNames,
  readFixture,
  readProducerFixture,
} from "../__testing__/docx";
import { documentFidelity } from "../editor/commands/fidelityQueries";
import { createEditorState } from "../editor/createEditor";
import { exportDocxReport } from "./exportDocx";
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

  it("reports a preserved block inside a cell through import, the editor and export", () => {
    const { doc, session, notes } = importDocx(
      makeDocx(
        `<w:p>${run("First")}</w:p>` +
          '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
          "<w:tr><w:tc><w:tcPr/><w:p><w:r>" +
          '<w:sym w:font="Wingdings" w:char="F0E0"/>' +
          "</w:r></w:p></w:tc></w:tr></w:tbl>"
      )
    );
    const expected = {
      severity: "placeholder",
      code: "preserved-block",
      part: "word/document.xml",
      block: 1,
      pos: 10,
      element: "w:p",
    };
    expect(notes).toEqual([expected]);

    const state = createEditorState(doc);
    expect(documentFidelity(state)).toEqual([{ ...expected, part: null }]);
    expect(exportDocxReport(state.doc, session).notes).toEqual([expected]);

    const edited = state.apply(state.tr.insertText("More ", 1));
    expect(documentFidelity(edited)).toEqual([
      { ...expected, part: null, pos: 15 },
    ]);
    expect(exportDocxReport(edited.doc, session).notes).toEqual([
      { ...expected, pos: 15 },
    ]);
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

describe("documentFidelity", () => {
  const BOOKMARKED =
    `<w:p>${run("a")}<w:bookmarkStart w:id="1" w:name="Here"/>` +
    '<w:bookmarkEnd w:id="1"/></w:p>';

  it("reads the open document, naming no part", () => {
    const state = createEditorState(importDocx(makeDocx(BOOKMARKED)).doc);

    expect(documentFidelity(state)).toEqual([
      {
        severity: "hidden",
        code: "range-marker",
        part: null,
        block: 0,
        pos: 2,
        element: "w:bookmarkStart",
      },
      {
        severity: "hidden",
        code: "range-marker",
        part: null,
        block: 0,
        pos: 3,
        element: "w:bookmarkEnd",
      },
    ]);
  });

  it("follows the document as it is edited", () => {
    const state = createEditorState(importDocx(makeDocx(BOOKMARKED)).doc);
    const edited = state.apply(state.tr.insertText("bcd", 2));

    expect(documentFidelity(edited).map((note) => note.pos)).toEqual([5, 6]);
  });
});

/**
 * What each fixture loses on the way in, held as a file a reviewer reads.
 *
 * A diff here is a change in what a document keeps, so it is approved deliberately rather than
 * accepted along with whatever else a PR touched. `docs/testing.md` says how to update them.
 */
describe("the fixture corpus", () => {
  it.each(fixtureNames)(
    "%s: opens with the recorded fidelity notes",
    async (name) => {
      const { notes } = importDocx(readFixture(name));

      await expect(`${JSON.stringify(notes, null, 2)}\n`).toMatchFileSnapshot(
        `./__snapshots__/fidelity/${name.replace(/\.docx$/, "")}.json`
      );
    }
  );

  /**
   * The producer lane, where the notes say what an editing session loses out of a document this
   * project did not write. Its snapshot is the reason the lane exists: nothing else says which
   * of a real word processor's constructs reach the reader as placeholders.
   */
  it.each(producerFixtureNames)(
    "%s: opens with the recorded fidelity notes",
    async (name) => {
      const { notes } = importDocx(readProducerFixture(name));

      await expect(`${JSON.stringify(notes, null, 2)}\n`).toMatchFileSnapshot(
        `./__snapshots__/fidelity/producers/${name.replace(/\.docx$/, "")}.json`
      );
    }
  );
});
