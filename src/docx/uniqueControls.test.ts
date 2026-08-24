// @vitest-environment node
import type { Mark, Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { docxSchema } from "../schema";
import { copiedControlPrefix } from "./sdt";
import { withUniqueControls } from "./uniqueControls";

const BOUND_PR =
  '<w:sdtPr><w:id w:val="7"/><w:dataBinding w:xpath="/contract/date"/></w:sdtPr>';
const PREFIX = `<w:sdt>${BOUND_PR}`;

/** What a copy of a bound control opens as: an id of its own and no binding left */
const COPY = /^<w:sdt><w:sdtPr><w:id w:val="\d+"\/><\/w:sdtPr>$/;

describe("the opening a copy of a control goes out as", () => {
  it("takes the id it is handed and leaves the binding behind", () => {
    expect(copiedControlPrefix(PREFIX, 42)).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="42"/></w:sdtPr>'
    );
  });

  it("keeps everything else the control said about itself", () => {
    const prefix =
      '<w:sdt w:rsidR="00A"><w:sdtPr><w:alias w:val="signedOn"/>' +
      '<w:lock w:val="sdtContentLocked"/></w:sdtPr>' +
      "<w:sdtEndPr><w:rPr/></w:sdtEndPr>";
    expect(copiedControlPrefix(prefix, 42)).toBe(
      '<w:sdt w:rsidR="00A"><w:sdtPr><w:alias w:val="signedOn"/>' +
        '<w:id w:val="42"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>' +
        "<w:sdtEndPr><w:rPr/></w:sdtEndPr>"
    );
  });

  it("hands back an opening it cannot make out just as it came", () => {
    expect(copiedControlPrefix("<w:sdt>", 42)).toBe("<w:sdt>");
  });
});

function control(sdtKey: number, sdtPrefix = PREFIX): Mark {
  return docxSchema.marks.sdt.create({ sdtPrefix, sdtKey });
}

const runMark = docxSchema.marks.run.create({ rPr: null });

function text(value: string, sdt: Mark | null): PMNode {
  return docxSchema.text(value, sdt ? [sdt, runMark] : [runMark]);
}

function paragraph(...inline: PMNode[]): PMNode {
  return docxSchema.nodes.paragraph.create(null, inline);
}

function doc(...blocks: PMNode[]): PMNode {
  return docxSchema.nodes.doc.create(null, blocks);
}

function cell(...blocks: PMNode[]): PMNode {
  return docxSchema.nodes.tableCell.create(null, blocks);
}

function table(...cells: PMNode[]): PMNode {
  return docxSchema.nodes.table.create(null, [
    docxSchema.nodes.tableRow.create(null, cells),
  ]);
}

/** The opening XML of every control the document holds, in the order they stand in */
function prefixesOf(node: PMNode): string[] {
  const found: string[] = [];
  node.descendants((child) => {
    const mark = child.marks.find((entry) => entry.type.name === "sdt");
    const prefix: unknown = mark?.attrs.sdtPrefix;
    if (typeof prefix === "string") found.push(prefix);
    return true;
  });
  return found;
}

describe("a control standing in more than one place", () => {
  it("keeps the original opening on the first of them and renames the rest", () => {
    const original = doc(
      paragraph(text("2026", control(3))),
      paragraph(text("-08-04", control(3)))
    );
    const next = withUniqueControls(original);
    const [first, second] = prefixesOf(next);

    expect(first).toBe(PREFIX);
    expect(second).toMatch(COPY);
  });

  it("leaves the paragraph the first of them stands in exactly as it was", () => {
    const original = doc(
      paragraph(text("2026", control(3))),
      paragraph(text("-08-04", control(3)))
    );
    const next = withUniqueControls(original);

    // An unedited block only goes out as its original XML while its node stays untouched
    expect(next.child(0)).toBe(original.child(0));
    expect(next.child(1)).not.toBe(original.child(1));
  });

  it("gives each copy after the first an opening of its own", () => {
    const next = withUniqueControls(
      doc(
        paragraph(text("2026", control(3))),
        paragraph(text("-08", control(3))),
        paragraph(text("-04", control(3)))
      )
    );
    expect(new Set(prefixesOf(next)).size).toBe(3);
  });

  it("counts a control broken in two inside the one paragraph", () => {
    const next = withUniqueControls(
      doc(
        paragraph(
          text("2026", control(3)),
          text(" AD ", null),
          text("08-04", control(3))
        )
      )
    );
    const [first, second] = prefixesOf(next);

    expect(first).toBe(PREFIX);
    expect(second).toMatch(COPY);
  });

  it("reaches the paragraphs inside a table as well", () => {
    const next = withUniqueControls(
      doc(
        table(
          cell(paragraph(text("2026", control(3)))),
          cell(paragraph(text("-08-04", control(3))))
        )
      )
    );
    const [first, second] = prefixesOf(next);

    expect(first).toBe(PREFIX);
    expect(second).toMatch(COPY);
  });

  it("the several runs of one control are that one control, not copies of it", () => {
    const original = doc(
      paragraph(text("2026", control(3)), text("AD", control(3)))
    );
    expect(withUniqueControls(original)).toBe(original);
  });
});

describe("controls that merely look alike", () => {
  it("are left as they came, opening XML and all", () => {
    const original = doc(
      paragraph(text("a", control(0)), text("b", control(1)))
    );
    expect(withUniqueControls(original)).toBe(original);
    expect(prefixesOf(original)).toEqual([PREFIX, PREFIX]);
  });
});
