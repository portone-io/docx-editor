// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { docxSchema } from "./index";
import { sameSource } from "./sourceEquality";

const paragraph = (attrs: Record<string, unknown>, text = "text") =>
  docxSchema.nodes.paragraph.create(attrs, [docxSchema.text(text)]);

const cell = (attrs: Record<string, unknown>) =>
  docxSchema.nodes.tableCell.create(attrs, [paragraph({})]);

const table = (
  attrs: Record<string, unknown>,
  first: Record<string, unknown>
) =>
  docxSchema.nodes.table.create(attrs, [
    docxSchema.nodes.tableRow.create(null, [cell(first), cell({})]),
  ]);

const run = (attrs: Record<string, unknown>) =>
  docxSchema.nodes.paragraph.create(null, [
    docxSchema.text("text", [docxSchema.marks.run.create(attrs)]),
  ]);

describe("judging two nodes by what they would be written as", () => {
  it("reads two paragraphs differing only in derived values as the same source", () => {
    const pPr = '<w:pPr><w:jc w:val="center"/></w:pPr>';
    const a = paragraph({ pPr, format: { align: "center" }, styleRun: null });
    const b = paragraph({ pPr, format: {}, styleRun: { bold: true } });

    expect(sameSource(a, b)).toBe(true);
    expect(a.eq(b)).toBe(false);
  });

  it("reads a run mark differing only in its derived format as the same source", () => {
    const rPr = "<w:rPr><w:b/></w:rPr>";
    const a = run({ rPr, format: { bold: true } });
    const b = run({ rPr, format: {} });

    expect(sameSource(a, b)).toBe(true);
    expect(a.eq(b)).toBe(false);
  });

  it("reads a run whose own XML changed as a different source", () => {
    expect(
      sameSource(
        run({ rPr: "<w:rPr><w:b/></w:rPr>" }),
        run({ rPr: "<w:rPr><w:i/></w:rPr>" })
      )
    ).toBe(false);
  });

  it("reads a cell whose own properties changed as a different source", () => {
    expect(
      sameSource(
        table({}, { tcPr: '<w:tcPr><w:shd w:fill="FF0000"/></w:tcPr>' }),
        table({}, { tcPr: '<w:tcPr><w:shd w:fill="00FF00"/></w:tcPr>' })
      )
    ).toBe(false);
  });

  it("reads a table whose cell format alone was re-derived as the same source", () => {
    const a = table({}, { format: { borders: { top: "single" } } });
    const b = table({}, { format: {} });

    expect(sameSource(a, b)).toBe(true);
    expect(a.eq(b)).toBe(false);
  });

  it("reads a cell whose lock was lifted as a different source", () => {
    expect(
      sameSource(
        table({}, { sdtContentsLocked: true }),
        table({}, { sdtContentsLocked: false })
      )
    ).toBe(false);
  });

  it("reads a note reference renumbered around an edit as the same reference", () => {
    const note = (label: string) =>
      docxSchema.nodes.noteReference.create({
        kind: "footnote",
        id: "2",
        label,
      });
    const a = note("1");
    const b = note("4");

    expect(sameSource(a, b)).toBe(true);
    expect(a.eq(b)).toBe(false);
  });

  it("reads a note reference of another id as a different source", () => {
    const note = (id: string) =>
      docxSchema.nodes.noteReference.create({ kind: "footnote", id });

    expect(sameSource(note("2"), note("3"))).toBe(false);
  });

  it("reads text that says something else as a different source", () => {
    expect(sameSource(paragraph({}, "one"), paragraph({}, "two"))).toBe(false);
  });

  it("reads a paragraph that gained a child as a different source", () => {
    const withBreak = docxSchema.nodes.paragraph.create(null, [
      docxSchema.text("text"),
      docxSchema.nodes.hardBreak.create({ brAttrs: null }),
    ]);

    expect(sameSource(paragraph({}), withBreak)).toBe(false);
  });

  it("reads a block that came from elsewhere in the file as a different source", () => {
    expect(
      sameSource(
        paragraph({ srcId: "opened:body:3" }),
        paragraph({ srcId: "opened:body:4" })
      )
    ).toBe(false);
  });

  /**
   * Every other case here hands both sides one attrs object, which `===` alone would settle. A
   * table's column widths are an array the writer writes from, and two tables opened apart hold
   * two arrays: the comparison has to read through them.
   */
  it("reads two tables built apart with the same column widths as the same source", () => {
    const a = table({ gridCols: [1000, 1000] }, {});
    const b = table({ gridCols: [1000, 1000] }, {});

    expect(a.attrs.gridCols).not.toBe(b.attrs.gridCols);
    expect(sameSource(a, b)).toBe(true);
  });

  it("reads two tables built apart with different column widths as a different source", () => {
    expect(
      sameSource(
        table({ gridCols: [1000, 1000] }, {}),
        table({ gridCols: [1000, 1200] }, {})
      )
    ).toBe(false);
    expect(
      sameSource(
        table({ gridCols: [1000, 1000] }, {}),
        table({ gridCols: [1000, 1000, 1000] }, {})
      )
    ).toBe(false);
  });

  it("reads a paragraph that lost its link as a different source", () => {
    const linked = docxSchema.nodes.paragraph.create(null, [
      docxSchema.text("text", [
        docxSchema.marks.link.create({ href: "https://example.com" }),
      ]),
    ]);

    expect(sameSource(paragraph({}), linked)).toBe(false);
  });
});
