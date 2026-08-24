// @vitest-environment jsdom
import { unzipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { bytesEqual, decode, makeDocx } from "../__testing__/docx";
import { createEditorState } from "../editor/createEditor";
import type { DocxExportError } from "../ooxml/errors";
import { docxSchema } from "../schema";
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

describe("body-level bookmarks", () => {
  it("imports a zero-length bookmark as invisible preservation nodes", () => {
    const { doc } = importDocx(makeDocx(ZERO_LENGTH));

    expect(doc.children.map((node) => node.type.name)).toEqual([
      "paragraph",
      "bookmarkBlock",
      "bookmarkBlock",
      "paragraph",
    ]);
    expect(doc.children.some((node) => node.type.name === "docxRaw")).toBe(
      false
    );
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
      if (start < 0 && node.type.name === "bookmarkBlock") start = pos;
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
      (node, index) => !(index === 0 && node.type.name === "bookmarkBlock")
    );
    const malformed = docxSchema.nodes.doc.create(null, blocks);

    expect(() => exportDocx(malformed, opened.session)).toThrowError(
      expect.objectContaining<Partial<DocxExportError>>({
        code: "malformed-xml",
      })
    );
  });
});
