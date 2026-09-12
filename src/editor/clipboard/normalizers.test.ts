// @vitest-environment jsdom

import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  makeDocx,
  makeNotesDocx,
  makeNumberedDocx,
  makeStyledDocx,
  makeStyledNumberedDocx,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { storyText } from "../../docx/story";
import { toParagraphFormat, toRunFormat } from "../../model/format";
import { NO_NEW_LISTS } from "../../numbering/listRegistry";
import { templateList } from "../../numbering/listTemplate";
import { docxSchema } from "../../schema";
import { type NoteKind, storyKey, storyNodeOf } from "../../schema/stories";
import { listRefOf } from "../commands/listCommands";
import { editorStateForSession } from "../createEditor";
import type { PastedContent } from "./htmlReader";
import { mapSliceNodes, normalizePasted } from "./normalizers";

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

/** One paragraph holding every shape of anchor a copied range may carry */
function anchoredParagraph(): PMNode {
  return paragraph({}, [
    docxSchema.nodes.commentStart.create({ id: "0" }),
    docxSchema.text("anchored"),
    docxSchema.nodes.commentEnd.create({ id: "0" }),
    docxSchema.nodes.commentReference.create({ id: "0" }),
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
}

/** A paragraph whose text calls a note, as a copy of one carries it */
function referring(kind: NoteKind, id: string): PMNode {
  return paragraph({}, [
    docxSchema.text("carried"),
    docxSchema.nodes.noteReference.create({
      id,
      kind,
      label: "1",
      referenceXml: `<w:${kind}Reference w:id="${id}"/>`,
    }),
  ]);
}

function noteStory(state: EditorState, kind: NoteKind, id: string): PMNode {
  const story = storyNodeOf(state.doc, storyKey(kind, id));
  if (story === null) throw new Error(`the document holds no ${kind} ${id}`);
  return story;
}

function idOf(node: PMNode | null | undefined): string {
  const id: unknown = node?.attrs.id;
  if (typeof id !== "string") throw new Error("the node names no id");
  return id;
}

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

  it("drops comment markers and bookmarks", () => {
    const state = stateOf(makeDocx(SOURCE));

    const normalized = normalizePasted(
      pasted(anchoredParagraph()),
      state,
      false
    );

    const first = blocks(normalized)[0];
    expect(first?.textContent).toBe("anchored");
    expect(first?.childCount).toBe(2);
    expect(first?.lastChild?.attrs.xml).toBe("<w:oMathPara/>");
  });

  /**
   * The mark a note's own body opens with stands for the number that note is called by, so in the
   * body it would draw a number the text never asked for.
   */
  it("drops a note's own reference mark pasted into the body", () => {
    const state = stateOf(makeDocx(SOURCE));
    const copied = paragraph({}, [
      docxSchema.nodes.rawRunContent.create({
        xml: "<w:footnoteRef/>",
        element: "footnoteRef",
        display: "chip",
      }),
      docxSchema.text("what the note said"),
      docxSchema.nodes.rawRunContent.create({
        xml: "<w:endnoteRef/>",
        element: "endnoteRef",
        display: "chip",
      }),
    ]);

    const normalized = normalizePasted(pasted(copied), state, false);

    const first = blocks(normalized)[0];
    expect(first?.textContent).toBe("what the note said");
    expect(first?.childCount).toBe(1);
  });

  it.each([
    ["footnote", "2"],
    ["endnote", "3"],
  ] as const)(
    "gives a pasted %s reference a new id and a copy of its note",
    (kind, id) => {
      const state = stateOf(makeNotesDocx());
      const story = noteStory(state, kind, id);

      const normalized = normalizePasted(
        {
          ...pasted(referring(kind, id)),
          noteStories: new Map([[storyKey(kind, id), story]]),
        },
        state,
        false
      );

      const reference = blocks(normalized)[0]?.lastChild;
      const pasted_id = idOf(reference);
      expect(pasted_id).not.toBe(id);
      // The XML it arrived as names the note it was copied from
      expect(reference?.attrs.referenceXml).toBeNull();
      const written = normalized.newStories?.get(storyKey(kind, pasted_id));
      expect(storyText(written ?? null)).toBe(storyText(story));
      // Written as an entry of its own rather than as the bytes of the original a second time
      expect(story.firstChild?.attrs.srcId).not.toBeNull();
      expect(written?.firstChild?.attrs.srcId).toBeNull();
    }
  );

  it("drops a footnote reference pasted from another editor", () => {
    const state = stateOf(makeNotesDocx());

    const normalized = normalizePasted(
      pasted(referring("footnote", "2")),
      state,
      false
    );

    const first = blocks(normalized)[0];
    expect(first?.textContent).toBe("carried");
    expect(first?.childCount).toBe(1);
    expect(normalized.newStories).toBeUndefined();
  });

  it("keeps a dragged footnote reference and its id on a move drop", () => {
    const state = stateOf(makeNotesDocx());

    const normalized = normalizePasted(
      pasted(referring("footnote", "2")),
      state,
      true
    );

    expect(idOf(blocks(normalized)[0]?.lastChild)).toBe("2");
    expect(normalized.newStories).toBeUndefined();
  });

  it("drops a permission and a move range the copied text stood in", () => {
    const state = stateOf(makeDocx(SOURCE));
    const ranged = paragraph({}, [
      docxSchema.nodes.rawInline.create({
        xml: '<w:permStart w:id="1" w:edGrp="everyone"/>',
        element: "permStart",
      }),
      docxSchema.text("ranged"),
      docxSchema.nodes.rawInline.create({
        xml: '<w:moveToRangeEnd w:id="4"/>',
        element: "moveToRangeEnd",
      }),
    ]);

    const normalized = normalizePasted(pasted(ranged), state, false);

    const first = blocks(normalized)[0];
    expect(first?.textContent).toBe("ranged");
    expect(first?.childCount).toBe(1);
  });

  it("keeps the anchors on a move drop, whose source goes as the drop lands", () => {
    const state = stateOf(makeDocx(SOURCE));
    const anchored = anchoredParagraph();

    const normalized = normalizePasted(pasted(anchored), state, true);

    // Detaching them would leave markers the document cannot write back, so the guard would
    // refuse the drag whole rather than move the range
    expect(blocks(normalized)[0]?.childCount).toBe(anchored.childCount);
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
