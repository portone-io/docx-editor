// @vitest-environment node
import type { Attrs, Mark, Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { exportErrorCode } from "../__testing__/docx";
import { docxSchema } from "../schema";
import { withUniqueIdentities } from "./identities";
import { copiedControlPrefix } from "./sdt";

const BOUND_PR =
  '<w:sdtPr><w:id w:val="7"/><w:dataBinding w:xpath="/contract/date"/></w:sdtPr>';
const PREFIX = `<w:sdt>${BOUND_PR}`;

/** What a copy of a bound control opens as: an id of its own and no binding left */
const COPY = /^<w:sdt><w:sdtPr><w:id w:val="\d+"\/><\/w:sdtPr>$/;

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

const SOURCE = "d1-abc:body:4";
const IDS = 'w14:paraId="1EADBEEF" w14:textId="77777777"';

/** A paragraph as it stands after being opened: named for its block and carrying its identifiers */
function opened(attrs: Attrs, ...inline: PMNode[]): PMNode {
  return docxSchema.nodes.paragraph.create(
    { srcId: SOURCE, pAttrs: IDS, ...attrs },
    inline
  );
}

describe("the source block rule", () => {
  it("the first claimant of a source block goes out verbatim and the second is rebuilt", () => {
    const original = doc(
      opened({}, text("source", null)),
      opened({}, text("source", null))
    );
    const next = withUniqueIdentities(original);

    expect(next.child(0)).toBe(original.child(0));
    expect(next.child(1).attrs.srcId).toBeNull();
    expect(next.child(1).textContent).toBe("source");
  });

  it.each(["docxRaw", "bookmarkBlock"])(
    "a preserved %s standing twice refuses export as unsupported-content",
    (name) => {
      const preserved = () =>
        docxSchema.nodes[name].create({ srcId: SOURCE, name: "w:tbl" });
      expect(
        exportErrorCode(() =>
          withUniqueIdentities(doc(preserved(), preserved()))
        )
      ).toBe("unsupported-content");
    }
  );

  it("names the block by the text of its key, so a number claims like a session key", () => {
    const original = doc(
      opened({ srcId: 4 }, text("a", null)),
      opened({ srcId: 4 }, text("b", null))
    );
    const next = withUniqueIdentities(original);

    expect(next.child(0)).toBe(original.child(0));
    expect(next.child(1).attrs.srcId).toBeNull();
  });
});

describe("the paragraph id rule", () => {
  it("a rebuilt second claimant carries no paraId or textId", () => {
    const original = doc(
      opened({ srcId: null }, text("a", null)),
      opened({ srcId: null }, text("b", null))
    );
    const next = withUniqueIdentities(original);

    expect(next.child(0)).toBe(original.child(0));
    expect(next.child(1).attrs.pAttrs).toBeNull();
  });

  it("keeps whatever else the opening tag said", () => {
    const next = withUniqueIdentities(
      doc(
        opened({ srcId: null }),
        opened({ srcId: null, pAttrs: `w:rsidR="00A1" ${IDS}` })
      )
    );
    expect(next.child(1).attrs.pAttrs).toBe('w:rsidR="00A1"');
  });

  it("reads the identifier as the number it spells, so two spellings of one value are one name", () => {
    const next = withUniqueIdentities(
      doc(
        opened({ srcId: null }),
        opened({ srcId: null, pAttrs: 'w14:paraId="1eadbeef"' })
      )
    );
    expect(next.child(1).attrs.pAttrs).toBeNull();
  });

  it.each(["00000000", "80000000", "1EADBEEF0", "DEADBEEG"])(
    "leaves a value no reader takes for an identifier (%s) as it came",
    (value) => {
      const original = doc(
        opened({ srcId: null, pAttrs: `w14:paraId="${value}"` }),
        opened({ srcId: null, pAttrs: `w14:paraId="${value}"` })
      );
      expect(withUniqueIdentities(original)).toBe(original);
    }
  );
});

describe("the rules together", () => {
  it("rules apply in order so a released block is also released of its paraId", () => {
    const original = doc(opened({}), opened({}));
    const next = withUniqueIdentities(original);

    expect(next.child(0)).toBe(original.child(0));
    expect(next.child(1).attrs).toMatchObject({ srcId: null, pAttrs: null });
  });

  it("a document with no duplicate is the very same node", () => {
    const original = doc(
      opened({}, text("2026", control(3))),
      opened({ srcId: "d1-abc:body:5", pAttrs: 'w14:paraId="00000042"' }),
      docxSchema.nodes.docxRaw.create({
        srcId: "d1-abc:body:6",
        name: "w:sectPr",
      }),
      table(cell(paragraph(text("-08-04", control(4)))))
    );
    expect(withUniqueIdentities(original)).toBe(original);
  });
});

describe("the control rule", () => {
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

  describe("a control standing in more than one place", () => {
    it("keeps the original opening on the first of them and renames the rest", () => {
      const original = doc(
        paragraph(text("2026", control(3))),
        paragraph(text("-08-04", control(3)))
      );
      const next = withUniqueIdentities(original);
      const [first, second] = prefixesOf(next);

      expect(first).toBe(PREFIX);
      expect(second).toMatch(COPY);
    });

    it("leaves the paragraph the first of them stands in exactly as it was", () => {
      const original = doc(
        paragraph(text("2026", control(3))),
        paragraph(text("-08-04", control(3)))
      );
      const next = withUniqueIdentities(original);

      // An unedited block only goes out as its original XML while its node stays untouched
      expect(next.child(0)).toBe(original.child(0));
      expect(next.child(1)).not.toBe(original.child(1));
    });

    it("gives each copy after the first an opening of its own", () => {
      const next = withUniqueIdentities(
        doc(
          paragraph(text("2026", control(3))),
          paragraph(text("-08", control(3))),
          paragraph(text("-04", control(3)))
        )
      );
      expect(new Set(prefixesOf(next)).size).toBe(3);
    });

    it("counts a control broken in two inside the one paragraph", () => {
      const next = withUniqueIdentities(
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
      const next = withUniqueIdentities(
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

    // A rule works on a block rather than on a node because of this: the runs of one control
    // are that control, and only a paragraph-wide walk can tell them from a copy
    it("the several runs of one control are that one control, not copies of it", () => {
      const original = doc(
        paragraph(text("2026", control(3)), text("AD", control(3)))
      );
      expect(withUniqueIdentities(original)).toBe(original);
    });
  });

  describe("controls that merely look alike", () => {
    it("are left as they came, opening XML and all", () => {
      const original = doc(
        paragraph(text("a", control(0)), text("b", control(1)))
      );
      expect(withUniqueIdentities(original)).toBe(original);
      expect(prefixesOf(original)).toEqual([PREFIX, PREFIX]);
    });
  });
});
