// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { docxSchema } from "../schema";
import { CLONE_POLICIES, splitParagraphAttrs } from "./cloning";
import { importDocx } from "./importDocx";

const SECT_PR = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

/** The first paragraph of a document opened from this body */
function paragraphOf(body: string) {
  const first = importDocx(makeDocx(body)).doc.child(0);
  if (first.type !== docxSchema.nodes.paragraph) {
    throw new Error(`the body opened as ${first.type.name}`);
  }
  return first;
}

/** A paragraph carrying both identifiers, a section break, and formatting beside them */
function sectionParagraph() {
  return paragraphOf(
    '<w:p w14:paraId="DEADBEEF" w14:textId="77777777">' +
      `<w:pPr><w:jc w:val="center"/>${SECT_PR}</w:pPr>` +
      '<w:r><w:t xml:space="preserve">text</w:t></w:r></w:p>'
  );
}

/** The first cell of a document whose only table holds one, wrapped in a locked content control */
function lockedCell() {
  const body =
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
    "<w:tr><w:sdt>" +
    '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>' +
    "<w:sdtContent>" +
    '<w:tc><w:tcPr><w:gridSpan w:val="1"/></w:tcPr>' +
    '<w:p><w:r><w:t xml:space="preserve">cell</w:t></w:r></w:p></w:tc>' +
    "</w:sdtContent></w:sdt></w:tr></w:tbl>";
  const found: ReturnType<typeof paragraphOf>[] = [];
  importDocx(makeDocx(body)).doc.descendants((node) => {
    if (found.length === 0 && node.type.spec.tableRole === "cell") {
      found.push(node);
    }
    return found.length === 0;
  });
  const cell = found[0];
  if (cell === undefined) throw new Error("the body opened without a cell");
  return cell;
}

describe("the attrs a paragraph made from another paragraph carries", () => {
  it("the after half of a split keeps the section break and loses the paragraph ids", () => {
    const { after } = splitParagraphAttrs(sectionParagraph());

    expect(after.pPr).toContain("<w:sectPr>");
    expect(after.pAttrs).toBeNull();
    // The block the original was opened from is not this paragraph's to claim
    expect(after.srcId).toBeUndefined();
  });

  it("the before half loses the section break and keeps its text formatting", () => {
    const paragraph = sectionParagraph();
    const { before } = splitParagraphAttrs(paragraph);

    expect(before.pPr).toBe('<w:pPr><w:jc w:val="center"/></w:pPr>');
    // The half standing where the original stood keeps its name and its block
    expect(before.pAttrs).toBe(paragraph.attrs.pAttrs);
    expect(before.srcId).toBe(paragraph.attrs.srcId);
    expect(before.format).toEqual(paragraph.attrs.format);
  });

  it("a copy loses ids, section break and srcId", () => {
    const copy = CLONE_POLICIES.paragraph.attrs(sectionParagraph(), "copy");

    expect(copy.pAttrs).toBeNull();
    expect(copy.pPr).toBe('<w:pPr><w:jc w:val="center"/></w:pPr>');
    expect(copy.srcId).toBeUndefined();
  });

  it("a pPr left empty becomes null", () => {
    const paragraph = paragraphOf(`<w:p><w:pPr>${SECT_PR}</w:pPr></w:p>`);

    expect(splitParagraphAttrs(paragraph).before.pPr).toBeNull();
  });

  it("a section break inside w:pPrChange is not a section break", () => {
    const pPr =
      '<w:pPr><w:pPrChange w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">' +
      `<w:pPr>${SECT_PR}</w:pPr></w:pPrChange></w:pPr>`;
    const paragraph = paragraphOf(`<w:p>${pPr}</w:p>`);

    // Left byte for byte as it arrived, the tracked change included
    expect(splitParagraphAttrs(paragraph).before.pPr).toBe(pPr);
  });

  it("keeps the attributes standing beside the ids it takes away", () => {
    const paragraph = paragraphOf(
      '<w:p w:rsidR="00AB12CD" w14:paraId="DEADBEEF" w14:textId="77777777"/>'
    );

    expect(splitParagraphAttrs(paragraph).after.pAttrs).toBe(
      'w:rsidR="00AB12CD"'
    );
  });
});

describe("the attrs a cell or a row made from another one carries", () => {
  it("a cell clone leaves the control and its locks behind", () => {
    const cell = lockedCell();
    expect(cell.attrs.sdtContentsLocked).toBe(true);

    const clone = CLONE_POLICIES.tableCell.attrs(cell, "copy");

    expect(clone.tcPr).toBe(cell.attrs.tcPr);
    expect(clone.sdtPrefix).toBeUndefined();
    expect(clone.sdtContentsLocked).toBeUndefined();
    expect(clone.sdtDeletionLocked).toBeUndefined();
    // Worked out again from the grid the cell lands in, never carried over
    expect(clone.colspan).toBeUndefined();
    expect(clone.rowspan).toBeUndefined();
    expect(clone.colwidth).toBeUndefined();
  });

  it("a row clone leaves the property exceptions behind", () => {
    const body =
      "<w:tbl>" +
      '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
      "<w:tr>" +
      '<w:tblPrEx><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPrEx>' +
      '<w:trPr><w:trHeight w:val="300"/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t xml:space="preserve">cell</w:t></w:r></w:p></w:tc>' +
      "</w:tr></w:tbl>";
    const rows: ReturnType<typeof paragraphOf>[] = [];
    importDocx(makeDocx(body)).doc.descendants((node) => {
      if (rows.length === 0 && node.type.spec.tableRole === "row") {
        rows.push(node);
      }
      return rows.length === 0;
    });
    const row = rows[0];
    if (row === undefined) throw new Error("the body opened without a row");
    expect(row.attrs.tblPrEx).not.toBeNull();

    const clone = CLONE_POLICIES.tableRow.attrs(row, "copy");

    expect(clone.trPr).toBe(row.attrs.trPr);
    expect(clone.tblPrEx).toBeUndefined();
  });
});
