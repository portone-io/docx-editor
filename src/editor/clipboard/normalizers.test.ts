// @vitest-environment jsdom

import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  makeDocx,
  makeNumberedDocx,
  makeStyledDocx,
  makeStyledNumberedDocx,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { toParagraphFormat, toRunFormat } from "../../model/format";
import { NO_NEW_LISTS } from "../../numbering/listRegistry";
import { templateList } from "../../numbering/listTemplate";
import { docxSchema } from "../../schema";
import { listRefOf } from "../commands/listCommands";
import { editorStateForSession } from "../createEditor";
import type { PastedContent } from "./htmlReader";
import { mapSliceMarks, mapSliceNodes, normalizePasted } from "./normalizers";

const HEADING_STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
  '<w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="DocumentHeading1">' +
  '<w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>';

const SOURCE = "<w:p><w:r><w:t>source</w:t></w:r></w:p>";

function stateOf(bytes: Uint8Array): EditorState {
  return editorStateForSession(importDocx(bytes));
}

function paragraph(
  attrs: Record<string, unknown>,
  content: readonly PMNode[] = [docxSchema.text("pasted")]
): PMNode {
  return docxSchema.nodes.paragraph.create(
    attrs,
    Fragment.fromArray([...content])
  );
}

function pasted(...blocks: readonly PMNode[]): PastedContent {
  return {
    slice: new Slice(Fragment.fromArray([...blocks]), 0, 0),
    newLists: NO_NEW_LISTS,
  };
}

const listPPr = (numId: number) =>
  "<w:pPr><w:numPr>" +
  '<w:ilvl w:val="0"/>' +
  `<w:numId w:val="${numId}"/>` +
  "</w:numPr></w:pPr>";

const listParagraph = (numId: number) =>
  paragraph({
    pPr: listPPr(numId),
    format: { numbering: { numId, ilvl: 0 } },
  });

function blocks(content: PastedContent): PMNode[] {
  const nodes: PMNode[] = [];
  content.slice.content.forEach((node) => {
    nodes.push(node);
  });
  return nodes;
}

describe("mapping over a slice", () => {
  it("drops the nodes mapped to null and leaves an open end standing", () => {
    const slice = new Slice(
      Fragment.fromArray([paragraph({ srcId: "gone" }), paragraph({})]),
      1,
      1
    );
    const mapped = mapSliceNodes(slice, (node) =>
      node.attrs.srcId === "gone" ? null : node
    );

    // The open ends say how the paste joins the text around it, which the mapping does not decide
    expect(mapped.openStart).toBe(1);
    expect(mapped.openEnd).toBe(1);
    expect(mapped.content.childCount).toBe(2);
    expect(mapped.content.firstChild?.textContent).toBe("");
  });

  it("removes a node standing under no open end", () => {
    const slice = new Slice(
      Fragment.fromArray([paragraph({ srcId: "gone" }), paragraph({})]),
      0,
      0
    );
    const mapped = mapSliceNodes(slice, (node) =>
      node.attrs.srcId === "gone" ? null : node
    );

    expect(mapped.content.childCount).toBe(1);
  });

  it("rewrites the attrs of one mark type and leaves the others alone", () => {
    const marked = docxSchema.text("marked", [
      docxSchema.marks.run.create({ rPr: "<w:rPr><w:b/></w:rPr>" }),
      docxSchema.marks.link.create({ href: "https://example.com" }),
    ]);
    const mapped = mapSliceMarks(
      new Slice(Fragment.from(paragraph({}, [marked])), 0, 0),
      docxSchema.marks.run,
      (mark) => ({ ...mark.attrs, rPr: null })
    );

    const text = mapped.content.firstChild?.firstChild;
    expect(
      docxSchema.marks.run.isInSet(text?.marks ?? [])?.attrs.rPr
    ).toBeNull();
    expect(docxSchema.marks.link.isInSet(text?.marks ?? [])?.attrs.href).toBe(
      "https://example.com"
    );
  });
});

describe("normalizing a pasted slice", () => {
  it("clears srcId so the block is rebuilt rather than written from another block's XML", () => {
    const state = stateOf(makeDocx(SOURCE));
    const srcId = state.doc.firstChild?.attrs.srcId;
    expect(typeof srcId).toBe("string");

    const normalized = normalizePasted(
      pasted(paragraph({ srcId })),
      state,
      false
    );

    expect(blocks(normalized)[0]?.attrs.srcId).toBeNull();
  });

  it("keeps srcId on a move drop, which carries the same block rather than a copy", () => {
    const state = stateOf(makeDocx(SOURCE));
    const srcId = state.doc.firstChild?.attrs.srcId;

    const normalized = normalizePasted(
      pasted(paragraph({ srcId })),
      state,
      true
    );

    expect(blocks(normalized)[0]?.attrs.srcId).toBe(srcId);
  });

  it("keeps a preserved block holding its own XML and drops one that only names a fragment", () => {
    const state = stateOf(makeDocx(SOURCE));
    const held = docxSchema.nodes.rawBlock.create({
      xml: "<w:tbl/>",
      name: "w:tbl",
    });
    const named = docxSchema.nodes.rawBlock.create({
      srcId: state.doc.firstChild?.attrs.srcId,
      name: "w:tbl",
    });

    const normalized = normalizePasted(pasted(held, named), state, false);

    expect(blocks(normalized)).toHaveLength(1);
    expect(blocks(normalized)[0]?.attrs.xml).toBe("<w:tbl/>");
  });

  it("gives a list number the document knows nothing about one that it defines", () => {
    const state = stateOf(makeNumberedDocx(SOURCE));

    const normalized = normalizePasted(
      pasted(listParagraph(99), listParagraph(99)),
      state,
      false
    );

    const [first, second] = blocks(normalized);
    expect(first?.attrs.pPr).toContain('<w:numId w:val="2"/>');
    // The two paragraphs shared a list before the paste and still share one after it
    expect(listRefOf(second as PMNode)).toEqual({ numId: 2, ilvl: 0 });
    expect(normalized.newLists.get(2)).toEqual(templateList("numbered"));
  });

  it("starts the list as the kind it was where the copy was made", () => {
    const state = stateOf(makeNumberedDocx(SOURCE));

    const normalized = normalizePasted(
      {
        slice: pasted(listParagraph(99)).slice,
        newLists: NO_NEW_LISTS,
        listKinds: new Map([[99, "bullet"]]),
      },
      state,
      false
    );

    expect(normalized.newLists.get(2)).toEqual(templateList("bullet"));
  });

  it("leaves a list number the document already answers for alone", () => {
    const state = stateOf(makeNumberedDocx(SOURCE));

    const normalized = normalizePasted(pasted(listParagraph(1)), state, false);

    expect(blocks(normalized)[0]?.attrs.pPr).toContain('<w:numId w:val="1"/>');
    expect(normalized.newLists.size).toBe(0);
  });

  it("takes the numbering off when the document has nowhere to define a list", () => {
    const state = stateOf(makeDocx(SOURCE));

    const normalized = normalizePasted(pasted(listParagraph(99)), state, false);

    const first = blocks(normalized)[0];
    expect(first?.attrs.pPr ?? "").not.toContain("numPr");
    expect(listRefOf(first as PMNode)).toBeNull();
  });

  it("drops comment markers, bookmarks and note references", () => {
    const state = stateOf(makeDocx(SOURCE));
    const anchored = paragraph({}, [
      docxSchema.nodes.commentStart.create({ id: "0" }),
      docxSchema.text("anchored"),
      docxSchema.nodes.commentEnd.create({ id: "0" }),
      docxSchema.nodes.commentReference.create({ id: "0" }),
      docxSchema.nodes.noteReference.create({ id: "2", label: "1" }),
      docxSchema.nodes.rawInline.create({
        xml: '<w:bookmarkStart w:id="1" w:name="mark"/>',
        element: "bookmarkStart",
      }),
      // The gate on raw XML cannot tell this one from any other element a paragraph may hold
      docxSchema.nodes.rawInline.create({
        xml: '<w:commentRangeStart w:id="7"/>',
        element: null,
      }),
      docxSchema.nodes.rawInline.create({
        xml: "<w:oMathPara/>",
        element: "oMathPara",
      }),
    ]);

    const normalized = normalizePasted(pasted(anchored), state, false);

    const first = blocks(normalized)[0];
    expect(first?.textContent).toBe("anchored");
    expect(first?.childCount).toBe(2);
    expect(first?.lastChild?.attrs.xml).toBe("<w:oMathPara/>");
  });

  it("works the display values out again against the destination's styles", () => {
    const state = stateOf(makeStyledDocx(SOURCE, HEADING_STYLES));
    const heading = paragraph(
      {
        pPr: '<w:pPr><w:pStyle w:val="DocumentHeading1"/></w:pPr>',
        format: null,
        styleRun: null,
      },
      [docxSchema.text("pasted", [docxSchema.marks.run.create({ rPr: null })])]
    );

    const normalized = normalizePasted(pasted(heading), state, false);

    const first = blocks(normalized)[0];
    expect(toRunFormat(first?.attrs.styleRun)).toMatchObject({
      bold: true,
      fontSizePt: 20,
    });
    expect(
      toRunFormat(first?.firstChild?.marks[0]?.attrs.format)
    ).toMatchObject({ bold: true, fontSizePt: 20 });
  });

  it("works the reissued list number into what the paragraph is drawn with", () => {
    const state = stateOf(makeStyledNumberedDocx(SOURCE, HEADING_STYLES));

    const normalized = normalizePasted(pasted(listParagraph(99)), state, false);

    // The number the export writes and the number the marker is drawn from are the same one
    expect(
      toParagraphFormat(blocks(normalized)[0]?.attrs.format)?.numbering
    ).toEqual({ numId: 2, ilvl: 0 });
  });
});
