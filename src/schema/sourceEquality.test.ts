// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { docxSchema } from "./index";
import { sameSource, withoutDisplayAttrs } from "./sourceEquality";

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
    const note = (label: string, text: string) =>
      docxSchema.nodes.noteReference.create({
        kind: "footnote",
        id: "2",
        label,
        text,
      });
    const a = note("1", "the body");
    const b = note("4", "another body");

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

  it("reads a paragraph that lost its link as a different source", () => {
    const linked = docxSchema.nodes.paragraph.create(null, [
      docxSchema.text("text", [
        docxSchema.marks.link.create({ href: "https://example.com" }),
      ]),
    ]);

    expect(sameSource(paragraph({}), linked)).toBe(false);
  });
});

describe("putting the derived values back where the schema starts them", () => {
  it("clears them on the node and on everything inside it", () => {
    const derived = table({ format: { width: 1 } }, { format: { top: 1 } });
    const cleared = withoutDisplayAttrs(derived);

    expect(cleared.attrs.format).toBeNull();
    expect(cleared.child(0).child(0).attrs.format).toBeNull();
    expect(cleared.eq(withoutDisplayAttrs(table({}, {})))).toBe(true);
  });

  it("leaves what the exporter writes from alone", () => {
    const pPr = '<w:pPr><w:jc w:val="center"/></w:pPr>';
    const cleared = withoutDisplayAttrs(
      paragraph({ pPr, format: { align: "center" } })
    );

    expect(cleared.attrs.pPr).toBe(pPr);
    expect(cleared.textContent).toBe("text");
  });

  it("clears them on a mark as well", () => {
    const cleared = withoutDisplayAttrs(
      run({ rPr: "<w:rPr><w:b/></w:rPr>", format: { bold: true } })
    );

    expect(cleared.child(0).marks[0].attrs.format).toBeNull();
    expect(cleared.child(0).marks[0].attrs.rPr).toBe("<w:rPr><w:b/></w:rPr>");
  });
});
