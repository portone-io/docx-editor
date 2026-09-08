// @vitest-environment jsdom

import { unzipSync } from "fflate";
import {
  DOMSerializer,
  DOMParser as ProseMirrorDOMParser,
} from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  decode,
  fixtureNames,
  makeDocx,
  producerFixtureNames,
  readFixture,
  readProducerFixture,
} from "../__testing__/docx";
import { documentFidelity } from "../editor/commands/fidelityQueries";
import { createEditorState } from "../editor/createEditor";
import { docxSchema } from "../schema";
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
  it("reports no original block number for a malformed key read from the DOM", () => {
    const { doc } = importDocx(
      makeDocx('<w:customXml w:uri="urn:example" w:element="example"/>')
    );
    const host = document.createElement("div");
    host.appendChild(
      DOMSerializer.fromSchema(docxSchema).serializeFragment(doc.content)
    );
    const placeholder = host.querySelector("[data-src]");
    if (placeholder === null)
      throw new Error("no preserved block was rendered");
    placeholder.setAttribute("data-src", `opened:body:${"9".repeat(400)}`);

    const parsed = ProseMirrorDOMParser.fromSchema(docxSchema).parse(host);
    expect(documentFidelity(createEditorState(parsed))).toEqual([
      expect.objectContaining({ block: null, element: "w:customXml" }),
    ]);
  });

  it("reports a run child the editor cannot draw, leaving the paragraph editable", () => {
    const { doc, notes } = importDocx(
      makeDocx(
        `<w:p>${run("First")}</w:p><w:p><w:r><w:sym w:char="F0E0"/></w:r></w:p>`
      )
    );

    expect(doc.child(1).type.name).toBe("paragraph");
    expect(notes).toEqual([
      {
        severity: "placeholder",
        code: "preserved-run-content",
        part: "word/document.xml",
        block: 1,
        pos: 8,
        element: "w:sym",
      },
    ]);
  });

  /**
   * Import no longer stands a paragraph down (`./importParagraph`), so the code is reached only by
   * a placeholder read back from the DOM, and it is what tells one demotion from another.
   */
  it("reports a placeholder standing for a paragraph as a demoted paragraph", () => {
    const doc = docxSchema.nodes.doc.create(null, [
      docxSchema.nodes.rawBlock.create({ srcId: "opened:body:4", name: "w:p" }),
    ]);

    expect(fidelityNotesOf(doc, null)).toEqual([
      {
        severity: "placeholder",
        code: "paragraph-demoted",
        part: null,
        block: 4,
        pos: 0,
        element: "w:p",
      },
    ]);
  });

  it("reports a demoted table as a placeholder note of its own", () => {
    // A content control around a whole row is a row this reader cannot take apart, and the two
    // levels with no node to keep a stranger in stand the table down (`./importPolicy`)
    const notes = notesOf(
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
        "<w:sdt><w:sdtPr/><w:sdtContent>" +
        `<w:tr><w:tc><w:p>${run("a")}</w:p></w:tc></w:tr>` +
        "</w:sdtContent></w:sdt></w:tbl>"
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

  it("reports nothing for a row-level bookmark, which rides on the table it marks", () => {
    expect(
      notesOf(
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
          '<w:tr><w:bookmarkStart w:id="1" w:name="b" w:colFirst="0" w:colLast="0"/>' +
          `<w:tc><w:p>${run("a")}</w:p></w:tc><w:bookmarkEnd w:id="1"/></w:tr></w:tbl>`
      )
    ).toEqual([]);
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

  it("reports a revision container as a placeholder the reader can see", () => {
    expect(
      notesOf(
        `<w:p>${run("a")}<w:ins w:id="1" w:author="x" w:date="2026-01-01T00:00:00Z">` +
          `${run("inserted")}</w:ins>${run("b")}</w:p>`
      )
    ).toEqual([
      expect.objectContaining({
        severity: "placeholder",
        code: "preserved-inline",
        block: 0,
        element: "w:ins",
      }),
    ]);
  });

  it("stops reporting a revision container once an edit has taken it away", () => {
    const { doc, session } = importDocx(
      makeDocx(
        `<w:p>${run("a")}<w:ins w:id="1" w:author="x" w:date="2026-01-01T00:00:00Z">` +
          `${run("inserted")}</w:ins>${run("b")}</w:p>`
      )
    );
    const state = createEditorState(doc);
    expect(documentFidelity(state)).toHaveLength(1);

    // Nothing guards a container that carries its content whole, so a selection over it deletes it
    const emptied = state.apply(state.tr.delete(1, state.doc.content.size - 1));

    expect(documentFidelity(emptied)).toEqual([]);
    const written = exportDocxReport(emptied.doc, session);
    expect(written.notes).toEqual([]);
    expect(
      decode(unzipSync(written.bytes)[session.mainPartPath])
    ).not.toContain("<w:ins");
  });

  it("reports a proofing mark, which nothing draws, as hidden preserved inline", () => {
    expect(
      notesOf(`<w:p><w:proofErr w:type="spellStart"/>${run("a")}</w:p>`)
    ).toEqual([
      expect.objectContaining({
        severity: "hidden",
        code: "preserved-inline",
        element: "w:proofErr",
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
          "<w:tr><w:tc><w:tcPr/>" +
          '<w:customXml w:element="clause"><w:p/></w:customXml>' +
          "</w:tc></w:tr></w:tbl>"
      )
    );
    const expected = {
      severity: "placeholder",
      code: "preserved-block",
      part: "word/document.xml",
      block: 1,
      pos: 10,
      element: "w:customXml",
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
