// @vitest-environment jsdom
import { unzipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { bytesEqual, decode, makeDocx } from "../__testing__/docx";
import { createEditorState } from "../editor/createEditor";
import type { DocxExportError } from "../ooxml/errors";
import { docxSchema } from "../schema";
import { withEditedFirst } from "./__testing__/blockEdits";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const ZERO_LENGTH =
  `<w:p>${run("Before")}</w:p>` +
  '<w:bookmarkStart w:id="7" w:name="Here"/>' +
  '<w:bookmarkEnd w:id="7"/>' +
  `<w:p>${run("After")}</w:p>`;

const BODY_RANGE =
  '<w:bookmarkStart w:id="8" w:name="Appendix"/>' +
  `<w:p>${run("First")}</w:p>` +
  `<w:p>${run("Second")}</w:p>` +
  '<w:bookmarkEnd w:id="8"/>';

/** A bookmark over a column, which stands under the row rather than inside any cell */
const COLUMN_RANGE =
  '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/>' +
  '<w:gridCol w:w="1000"/></w:tblGrid><w:tr>' +
  '<w:bookmarkStart w:id="6" w:name="Column" w:colFirst="0" w:colLast="0"/>' +
  `<w:tc><w:p>${run("First")}</w:p></w:tc>` +
  `<w:tc><w:p>${run("Second")}</w:p></w:tc>` +
  '<w:bookmarkEnd w:id="6"/></w:tr></w:tbl>';

const MIXED_RANGE =
  '<w:bookmarkStart w:id="9" w:name="Mixed"/>' +
  `<w:p>${run("Inside")}<w:bookmarkEnd w:id="9"/></w:p>`;

function editParagraph(doc: PMNode, index: number, text: string): PMNode {
  let paragraphIndex = 0;
  const blocks: PMNode[] = [];
  doc.forEach((block) => {
    if (block.type.name !== "paragraph") {
      blocks.push(block);
      return;
    }
    if (paragraphIndex !== index) {
      blocks.push(block);
      paragraphIndex += 1;
      return;
    }
    const inline: PMNode[] = [];
    block.forEach((child) => {
      inline.push(child.isText ? docxSchema.text(text, child.marks) : child);
    });
    blocks.push(block.copy(Fragment.from(inline)));
    paragraphIndex += 1;
  });
  return docxSchema.nodes.doc.create(null, blocks);
}

function documentXml(bytes: Uint8Array): string {
  return decode(unzipSync(bytes)["word/document.xml"]);
}

describe("a bookmark over a column", () => {
  it("opens as a table rather than standing the table down", () => {
    const { doc } = importDocx(makeDocx(COLUMN_RANGE));

    expect(doc.child(0).type.name).toBe("table");
    expect(doc.child(0).textContent).toBe("FirstSecond");
  });

  it("keeps the untouched table byte-identical", () => {
    const bytes = makeDocx(COLUMN_RANGE);
    const opened = importDocx(bytes);
    const before = unzipSync(bytes)["word/document.xml"];
    const after = unzipSync(exportDocx(opened.doc, opened.session))[
      "word/document.xml"
    ];

    expect(bytesEqual(after, before)).toBe(true);
  });

  it("keeps both markers around the same cells once the table is rebuilt", () => {
    const opened = importDocx(makeDocx(COLUMN_RANGE));
    const xml = documentXml(
      exportDocx(withEditedFirst(opened.doc, "table", "Edited"), opened.session)
    );

    expect(xml).toMatch(
      /<w:tr><w:bookmarkStart w:id="6" .*Edited.*Second.*<w:bookmarkEnd w:id="6"\/><\/w:tr>/
    );
  });
});

describe("body-level bookmarks", () => {
  it("imports a zero-length bookmark as invisible preservation nodes", () => {
    const { doc } = importDocx(makeDocx(ZERO_LENGTH));

    expect(doc.children.map((node) => node.type.name)).toEqual([
      "paragraph",
      "rawBlock",
      "rawBlock",
      "paragraph",
    ]);
    // A marker holds a spot and draws nothing, which is what tells it from the placeholder a
    // block nobody could read stands as
    expect(
      doc.children
        .slice(1, 3)
        .map((node) => [node.attrs.display, node.attrs.guarded])
    ).toEqual([
      ["hidden", true],
      ["hidden", true],
    ]);
  });

  it("keeps untouched bookmark XML byte-identical", () => {
    const bytes = makeDocx(BODY_RANGE);
    const opened = importDocx(bytes);
    const before = unzipSync(bytes)["word/document.xml"];
    const after = unzipSync(exportDocx(opened.doc, opened.session))[
      "word/document.xml"
    ];

    expect(bytesEqual(after, before)).toBe(true);
  });

  it("keeps a zero-length bookmark between the same surrounding paragraphs", () => {
    const opened = importDocx(makeDocx(ZERO_LENGTH));
    const xml = documentXml(
      exportDocx(editParagraph(opened.doc, 1, "Edited after"), opened.session)
    );

    expect(xml).toMatch(
      /Before.*<w:bookmarkStart w:id="7" w:name="Here"\/><w:bookmarkEnd w:id="7"\/>.*Edited after/
    );
  });

  it("keeps a cross-paragraph range around the same paragraphs", () => {
    const opened = importDocx(makeDocx(BODY_RANGE));
    const xml = documentXml(
      exportDocx(editParagraph(opened.doc, 0, "Edited first"), opened.session)
    );

    expect(xml).toMatch(
      /<w:bookmarkStart w:id="8" w:name="Appendix"\/>.*Edited first.*Second.*<w:bookmarkEnd w:id="8"\/>/
    );
  });

  it("keeps a range whose markers cross body and paragraph contexts", () => {
    const opened = importDocx(makeDocx(MIXED_RANGE));
    const xml = documentXml(
      exportDocx(editParagraph(opened.doc, 0, "Edited"), opened.session)
    );

    expect(xml).toMatch(
      /<w:bookmarkStart w:id="9" w:name="Mixed"\/>.*Edited.*<w:bookmarkEnd w:id="9"\/>/
    );
  });

  it("refuses an editor transaction that removes only one marker", () => {
    const { doc } = importDocx(makeDocx(BODY_RANGE));
    const state = createEditorState(doc);
    let start = -1;
    state.doc.descendants((node, pos) => {
      if (start < 0 && node.attrs.guarded === true) start = pos;
      return start < 0;
    });
    if (start < 0) throw new Error("no bookmark marker in the test document");

    expect(
      state.apply(state.tr.delete(start, start + 1)).doc.eq(state.doc)
    ).toBe(true);
  });

  it("refuses to export a programmatic transform with an unmatched marker", () => {
    const opened = importDocx(makeDocx(BODY_RANGE));
    const blocks = opened.doc.children.filter(
      (node, index) => !(index === 0 && node.type.name === "rawBlock")
    );
    const malformed = docxSchema.nodes.doc.create(null, blocks);

    expect(() => exportDocx(malformed, opened.session)).toThrowError(
      expect.objectContaining<Partial<DocxExportError>>({
        code: "malformed-xml",
      })
    );
  });
});
