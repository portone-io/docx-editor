// @vitest-environment jsdom
/**
 * The block content control read as a container, and everything it has to keep on the way back.
 *
 * The cases are built from small bodies rather than read off the fixture, so what each one proves
 * stands in the test itself; `content-controls.docx` is what holds them all in one package and is
 * read here for the byte identity and the edited round trip.
 */
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  readFixture,
  withBlocks,
} from "../__testing__/docx";
import { docxSchema } from "../schema";
import { firstBlockIndex, withEditedBlock } from "./__testing__/blockEdits";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";

const FIXTURE = "content-controls.docx";

const P = (text: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** A control, written the way the fixture writes one: an id, then an optional lock, then its type */
function sdt(
  content: string,
  { id = 1, lock = "", type = "<w:richText/>" } = {}
): string {
  const locked = lock === "" ? "" : `<w:lock w:val="${lock}"/>`;
  return (
    `<w:sdt><w:sdtPr><w:id w:val="${id}"/>${locked}${type}</w:sdtPr>` +
    `<w:sdtContent>${content}</w:sdtContent></w:sdt>`
  );
}

function bodyOf(xml: string): PMNode {
  return importDocx(makeDocx(xml)).doc;
}

/** The type of every top-level block, which is what says whether a control opened as a container */
function blockTypes(doc: PMNode): string[] {
  const names: string[] = [];
  doc.forEach((block) => {
    names.push(block.type.name);
  });
  return names;
}

function onlyBlock(xml: string): PMNode {
  const doc = bodyOf(xml);
  expect(doc.childCount).toBe(1);
  return doc.child(0);
}

describe("a block control read as a container", () => {
  it("holds the blocks its sdtContent held", () => {
    const control = onlyBlock(sdt(P("First") + P("Second")));

    expect(control.type.name).toBe("sdtBlock");
    expect(blockTypes(control)).toEqual(["paragraph", "paragraph"]);
    expect(control.textContent).toBe("FirstSecond");
  });

  it("keeps the opening tag verbatim and names the block it was opened from", () => {
    const control = onlyBlock(sdt(P("First"), { id: 42 }));

    expect(control.attrs.sdtPrefix).toBe(
      '<w:sdt><w:sdtPr><w:id w:val="42"/><w:richText/></w:sdtPr>'
    );
    expect(typeof control.attrs.srcId).toBe("string");
  });

  it("gives its children no source of their own, as a table cell does", () => {
    const control = onlyBlock(sdt(P("First") + P("Second")));

    expect(control.child(0).attrs.srcId).toBeNull();
    expect(control.child(1).attrs.srcId).toBeNull();
  });

  it("opens a control carrying no type at all, which is a rich text control", () => {
    const control = onlyBlock(sdt(P("First"), { type: "" }));

    expect(control.type.name).toBe("sdtBlock");
  });

  it("counts one control apart from the next, inline controls included", () => {
    const doc = bodyOf(
      sdt(P("First"), { id: 1 }) + sdt(P("Second"), { id: 2 })
    );

    expect(doc.child(0).attrs.key).not.toBe(doc.child(1).attrs.key);
  });

  it("reads both clauses of a lock off the control", () => {
    const shut = onlyBlock(sdt(P("First"), { lock: "sdtContentLocked" }));
    const undeletable = onlyBlock(sdt(P("First"), { lock: "sdtLocked" }));

    expect(shut.attrs).toMatchObject({
      contentsLocked: true,
      deletionLocked: true,
    });
    expect(undeletable.attrs).toMatchObject({
      contentsLocked: false,
      deletionLocked: true,
    });
  });
});

describe("what a block control may hold", () => {
  it("holds a table", () => {
    const control = onlyBlock(
      sdt(`<w:tbl><w:tr><w:tc>${P("Cell")}</w:tc></w:tr></w:tbl>`)
    );

    expect(blockTypes(control)).toEqual(["table"]);
  });

  it("holds another control, however the two are typed", () => {
    const control = onlyBlock(
      sdt(sdt(P("Inner"), { id: 2 }), { id: 1, type: "<w:group/>" })
    );

    expect(blockTypes(control)).toEqual(["sdtBlock"]);
    expect(control.child(0).textContent).toBe("Inner");
  });

  it("keeps a block it has no model for as a placeholder inside itself", () => {
    const control = onlyBlock(
      sdt(P("First") + '<w:customXml w:uri="urn:x" w:element="y"/>')
    );

    expect(blockTypes(control)).toEqual(["paragraph", "rawBlock"]);
  });

  it("stands inside a table cell, around the blocks of that cell", () => {
    const doc = bodyOf(
      `<w:tbl><w:tr><w:tc>${sdt(P("Inside"))}</w:tc></w:tr></w:tbl>`
    );
    const cell = doc.child(0).child(0).child(0);

    expect(blockTypes(cell)).toEqual(["sdtBlock"]);
    expect(cell.child(0).textContent).toBe("Inside");
  });
});

describe("a block control kept whole instead", () => {
  it.each([
    ['<w:text w:multiLine="0"/>'],
    ["<w:picture/>"],
    ["<w:date/>"],
    ["<w:comboBox/>"],
    ["<w:dropDownList/>"],
    ["<w14:checkbox/>"],
  ])("keeps %s, a type whose content the specification restrains", (type) => {
    expect(blockTypes(bodyOf(sdt(P("First"), { type })))).toEqual(["rawBlock"]);
  });

  it("reads a restrained type by its namespace, not by its local name alone", () => {
    expect(
      blockTypes(bodyOf(sdt(P("First"), { type: "<w:checkbox/>" })))
    ).toEqual(["sdtBlock"]);
  });

  it("keeps a control holding nothing, since a container cannot say nothing", () => {
    expect(blockTypes(bodyOf(sdt("")))).toEqual(["rawBlock"]);
    expect(
      blockTypes(bodyOf('<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr></w:sdt>'))
    ).toEqual(["rawBlock"]);
  });

  it("keeps a control whose wrapper this editor could not write back", () => {
    const attributed =
      '<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr>' +
      `<w:sdtContent w:x="1">${P("First")}</w:sdtContent></w:sdt>`;

    expect(blockTypes(bodyOf(attributed))).toEqual(["rawBlock"]);
  });
});

describe("a block control written back", () => {
  /**
   * `roundtrip.test.ts` holds the untouched identity over every fixture. What this adds is the
   * identity of a control an edit stood next to: the whole point of reading one as a container is
   * that the rest of the document keeps writing back as it was, and so does the control.
   */
  it("keeps its own bytes while an edit rewrites the block beside it", () => {
    const bytes = readFixture(FIXTURE);
    const { doc, session } = importDocx(bytes);
    const original = documentXmlOf(doc, session);
    const controls = Array.from(original.matchAll(/<w:sdt>.*?<\/w:sdt>/gs));
    const index = firstBlockIndex(doc, "paragraph");

    const out = documentXmlOf(withEditedBlock(doc, index, "Edited"), session);

    expect(controls.length).toBeGreaterThan(0);
    for (const [control] of controls) expect(out).toContain(control);
    expect(out).toContain("Edited");
  });

  it("writes the shape the reader takes back apart once it was edited", () => {
    const bytes = makeDocx(sdt(P("First") + P("Second")));
    const { doc, session } = importDocx(bytes);
    const control = doc.child(0);
    const edited = control.copy(
      Fragment.from([
        control.child(0).copy(Fragment.from(docxSchema.text("Changed"))),
        control.child(1),
      ])
    );

    const written = exportDocx(withBlocks(doc, [edited]), session);

    expect(documentXmlOf(withBlocks(doc, [edited]), session)).toContain(
      `<w:sdtContent>${P("Changed")}${P("Second")}</w:sdtContent></w:sdt>`
    );
    expect(blockTypes(importDocx(written).doc)).toEqual(["sdtBlock"]);
  });

  it("re-reads an edited control as the same shape, its lock and its id included", () => {
    const bytes = makeDocx(
      sdt(P("First") + P("Second"), { id: 42, lock: "sdtContentLocked" })
    );
    const { doc, session } = importDocx(bytes);
    const control = doc.child(0);
    const edited = control.copy(
      Fragment.from([
        control.child(0).copy(Fragment.from(docxSchema.text("Changed"))),
        control.child(1),
      ])
    );

    const reread = importDocx(
      exportDocx(withBlocks(doc, [edited]), session)
    ).doc.child(0);

    expect(reread.attrs.sdtPrefix).toBe(control.attrs.sdtPrefix);
    expect(reread.attrs.contentsLocked).toBe(true);
    expect(reread.textContent).toBe("ChangedSecond");
  });

  it("gives a control that goes out twice a name of its own", () => {
    const bytes = makeDocx(sdt(P("First"), { id: 42 }));
    const { doc, session } = importDocx(bytes);
    const control = doc.child(0);

    const out = documentXmlOf(withBlocks(doc, [control, control]), session);
    const ids = Array.from(
      out.matchAll(/<w:id w:val="(\d+)"\/>/g),
      (m) => m[1]
    );

    expect(ids).toHaveLength(2);
    expect(ids[0]).toBe("42");
    expect(ids[1]).not.toBe("42");
  });
});
