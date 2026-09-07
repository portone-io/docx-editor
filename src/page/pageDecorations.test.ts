// @vitest-environment jsdom
import { redo, redoDepth, undo, undoDepth } from "prosemirror-history";
import type { Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it, vi } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { serializeParagraph } from "../docx/serializeParagraph";
import { createEditorState } from "../editor/createEditor";
import { docxSchema, isPageBreak } from "../schema";
import { editorAttributes } from "../styles/classNames";
import { setPageMarks } from "./pageDecorations";

let view: EditorView | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
  document.body.replaceChildren();
});

function mounted(doc: PMNode): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  view = new EditorView(mount, { state: createEditorState(doc) });
  return view;
}

/** One editor with two paragraphs. The first carries a gap of its own above it */
function editor(): EditorView {
  return mounted(
    docxSchema.nodes.doc.create(null, [
      docxSchema.nodes.paragraph.create(
        {
          pPr: '<w:pPr><w:spacing w:before="240"/></w:pPr>',
          format: { spaceBeforePt: 12 },
        },
        [docxSchema.text("first paragraph")]
      ),
      docxSchema.nodes.paragraph.create({}, [
        docxSchema.text("second paragraph"),
      ]),
    ])
  );
}

/** The position where the second block starts */
function secondBlock(live: EditorView): number {
  return live.state.doc.child(0).nodeSize;
}

function paragraphs(live: EditorView): HTMLElement[] {
  return Array.from(live.dom.querySelectorAll("p"));
}

/**
 * The pushed value. It layers on afterwards rather than overwriting the gap the paragraph
 * already had
 */
function pushed(element: HTMLElement | undefined): string {
  return element?.style.getPropertyValue("margin-block-start") ?? "";
}

/** The gap the paragraph already had of its own */
function own(element: HTMLElement | undefined): string {
  return element?.style.marginTop ?? "";
}

/** One editor with a single paragraph broken twice, so the breaks have ordinals that can shift */
function brokenEditor(): EditorView {
  return mounted(
    docxSchema.nodes.doc.create(null, [
      docxSchema.nodes.paragraph.create({}, [
        docxSchema.text("aaa"),
        pageBreak(),
        docxSchema.text("bbb"),
        pageBreak(),
        docxSchema.text("ccc"),
      ]),
    ])
  );
}

function pageBreak(): PMNode {
  return docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' });
}

/** The height standing on each break's space, in document order */
function spaceHeights(live: EditorView): string[] {
  return Array.from(
    live.dom.querySelectorAll(`[${editorAttributes.pageBreakSpace}]`),
    (element) => element.getAttribute(editorAttributes.pageBreakSpace) ?? ""
  );
}

/** Where the document's page breaks stand, in document order */
function breaks(live: EditorView): number[] {
  const found: number[] = [];
  live.state.doc.descendants((node, pos) => {
    if (node.type === docxSchema.nodes.hardBreak) found.push(pos);
  });
  return found;
}

/** Where the first `br` of the document stands */
function firstBreak(live: EditorView): number {
  const [first] = breaks(live);
  if (first === undefined) throw new Error("the document holds no break");
  return first;
}

/** The two heights the layout works out for the two breaks */
function twoSpaces(live: EditorView) {
  const heights = [111, 222];
  return breaks(live).map((at, index) => ({
    at,
    height: heights[index] ?? 0,
  }));
}

function tableEditor(): {
  live: EditorView;
  secondRow: number;
  headerRow: number;
} {
  const first = docxSchema.nodes.tableRow.create(
    { format: { repeatHeader: true } },
    [cell("Heading A"), cell("Heading B")]
  );
  const second = docxSchema.nodes.tableRow.create({}, [cell("A"), cell("B")]);
  const table = docxSchema.nodes.table.create({ gridCols: [1000, 1000] }, [
    first,
    second,
  ]);
  return {
    live: mounted(docxSchema.nodes.doc.create(null, [table])),
    headerRow: 1,
    secondRow: 1 + first.nodeSize,
  };
}

function cell(text: string) {
  return docxSchema.nodes.tableCell.create({}, [
    docxSchema.nodes.paragraph.create({}, [docxSchema.text(text)]),
  ]);
}

describe("setPageMarks", () => {
  it("updates the measured push even when the total margin stays the same", () => {
    const live = editor();
    setPageMarks(live, {
      pushes: [{ pos: 0, marginTop: 40, push: 24 }],
      cuts: [],
    });
    setPageMarks(live, {
      pushes: [{ pos: 0, marginTop: 40, push: 16 }],
      cuts: [],
    });
    expect(paragraphs(live)[0]?.getAttribute(editorAttributes.pagePush)).toBe(
      "16"
    );
  });

  it("keeps a cut when its row only changes markup", () => {
    const { live, secondRow } = tableEditor();
    setPageMarks(live, { pushes: [], cuts: [{ at: secondRow, height: 240 }] });
    const row = live.state.doc.nodeAt(secondRow);
    live.dispatch(
      live.state.tr.setNodeMarkup(secondRow, undefined, {
        ...row?.attrs,
        format: { cantSplit: true },
      })
    );
    expect(
      live.dom
        .querySelector(`[${editorAttributes.tablePageSpace}]`)
        ?.getAttribute(editorAttributes.tablePageSpace)
    ).toBe("240");
  });

  it("keeps a cut with its original row after a row is inserted before it", () => {
    const { live, secondRow } = tableEditor();
    setPageMarks(live, { pushes: [], cuts: [{ at: secondRow, height: 240 }] });
    live.dispatch(
      live.state.tr.insert(
        secondRow,
        docxSchema.nodes.tableRow.create({}, [cell("inserted"), cell("row")])
      )
    );
    const spacer = live.dom.querySelector(
      `[${editorAttributes.tablePageSpace}]`
    );
    expect(spacer?.previousElementSibling?.textContent).toBe("insertedrow");
    expect(spacer?.nextElementSibling?.nextElementSibling?.textContent).toBe(
      "AB"
    );
  });

  it("remeasures a table moved by deletion and insertion instead of transplanting its old cut", () => {
    const { live, secondRow } = tableEditor();
    const table = live.state.doc.child(0);
    const paragraph = docxSchema.nodes.paragraph.create({}, [
      docxSchema.text("before"),
    ]);
    live.dispatch(live.state.tr.insert(table.nodeSize, paragraph));
    setPageMarks(live, { pushes: [], cuts: [{ at: secondRow, height: 240 }] });

    live.dispatch(
      live.state.tr.delete(0, table.nodeSize).insert(paragraph.nodeSize, table)
    );
    expect(
      live.dom.querySelector(`[${editorAttributes.tablePageSpace}]`)
    ).toBeNull();
    expect(live.state.doc.child(1).eq(table)).toBe(true);

    setPageMarks(live, {
      pushes: [],
      cuts: [{ at: paragraph.nodeSize + secondRow, height: 120 }],
    });
    expect(
      live.dom
        .querySelector(`[${editorAttributes.tablePageSpace}]`)
        ?.getAttribute(editorAttributes.tablePageSpace)
    ).toBe("120");
  });

  it("only the pushed block gets a wider gap", () => {
    const live = editor();
    setPageMarks(live, {
      pushes: [{ pos: secondBlock(live), marginTop: 300, push: 300 }],
      cuts: [],
    });

    const [first, second] = paragraphs(live);
    expect(pushed(first)).toBe("");
    expect(pushed(second)).toBe("300px");
    expect(second?.getAttribute("data-page-push")).toBe("300");
  });

  it("the gap the paragraph already had of its own remains after the push is taken away", () => {
    const live = editor();
    setPageMarks(live, {
      pushes: [{ pos: 0, marginTop: 316, push: 300 }],
      cuts: [],
    });
    expect(pushed(paragraphs(live)[0])).toBe("316px");
    expect(own(paragraphs(live)[0])).toBe("12pt");

    setPageMarks(live, { pushes: [], cuts: [] });
    expect(pushed(paragraphs(live)[0])).toBe("");
    expect(own(paragraphs(live)[0])).toBe("12pt");
    expect(paragraphs(live)[0]?.hasAttribute("data-page-push")).toBe(false);
  });

  it("does not touch the document when the same value is applied again", () => {
    const live = editor();
    const marks = { pushes: [{ pos: 0, marginTop: 16, push: 16 }], cuts: [] };
    setPageMarks(live, marks);
    const before = live.state;
    setPageMarks(live, marks);
    expect(live.state).toBe(before);
  });

  it("a push leaves nothing behind in the document or in the edit history", () => {
    const live = editor();
    const doc = live.state.doc;
    setPageMarks(live, {
      pushes: [{ pos: 0, marginTop: 40, push: 24 }],
      cuts: [],
    });
    expect(live.state.doc).toBe(doc);
  });

  it("forgets a cut when a page break becomes a line break", () => {
    const live = brokenEditor();
    const at = firstBreak(live);
    setPageMarks(live, { pushes: [], cuts: twoSpaces(live) });

    live.dispatch(live.state.tr.setNodeAttribute(at, "brAttrs", null));
    expect(spaceHeights(live)).toEqual(["222"]);
    // Restoring the break before a new measurement must not resurrect its stale height.
    live.dispatch(
      live.state.tr.setNodeAttribute(at, "brAttrs", 'w:type="page"')
    );
    expect(spaceHeights(live)).toEqual(["0", "222"]);
  });

  it("opens each break the height it was given", () => {
    const live = brokenEditor();
    setPageMarks(live, { pushes: [], cuts: twoSpaces(live) });
    expect(spaceHeights(live)).toEqual(["111", "222"]);
  });

  it("every break is left with an empty space when they are taken away", () => {
    const live = brokenEditor();
    setPageMarks(live, { pushes: [], cuts: twoSpaces(live) });

    setPageMarks(live, { pushes: [], cuts: [] });
    expect(spaceHeights(live)).toEqual(["0", "0"]);
  });

  it("does not touch the document when the same values are applied again", () => {
    const live = brokenEditor();
    setPageMarks(live, { pushes: [], cuts: twoSpaces(live) });
    const before = live.state;
    setPageMarks(live, { pushes: [], cuts: twoSpaces(live) });
    expect(live.state).toBe(before);
  });

  it("a space leaves nothing behind in the document or in the edit history", () => {
    const live = brokenEditor();
    const doc = live.state.doc;
    setPageMarks(live, { pushes: [], cuts: twoSpaces(live) });
    expect(live.state.doc).toBe(doc);
  });

  /**
   * A space is measured for one break and would be a page-sized height on any other, so it has to
   * follow the very `br` it was worked out for. Keyed by the ordinal the layout counted, the
   * survivor of a deletion took the height of the break that went away.
   */
  it("a space stays with its own break when an earlier break is deleted", () => {
    const live = brokenEditor();
    setPageMarks(live, { pushes: [], cuts: twoSpaces(live) });

    const first = firstBreak(live);
    live.dispatch(live.state.tr.delete(first, first + 1));

    expect(spaceHeights(live)).toEqual(["222"]);
  });

  it("a break put in before another one starts out with no space of its own", () => {
    const live = brokenEditor();
    setPageMarks(live, { pushes: [], cuts: twoSpaces(live) });

    live.dispatch(live.state.tr.insert(firstBreak(live), pageBreak()));

    expect(spaceHeights(live)).toEqual(["0", "111", "222"]);
  });

  it("inserts a non-editable spacer and header projection before the continued row", () => {
    const { live, secondRow } = tableEditor();
    setPageMarks(live, {
      pushes: [],
      cuts: [{ at: secondRow, height: 240 }],
    });

    const rows = Array.from(live.dom.querySelectorAll("tr"));
    expect(rows.map((row) => row.textContent)).toEqual([
      "Heading AHeading B",
      "",
      "Heading AHeading B",
      "AB",
    ]);
    expect(rows[1]?.getAttribute(editorAttributes.tablePageSpace)).toBe("240");
    expect(rows[2]?.hasAttribute(editorAttributes.tableRepeatedHeader)).toBe(
      true
    );
    expect(rows[2]?.getAttribute("contenteditable")).toBe("false");
    expect(rows[2]?.getAttribute("aria-hidden")).toBe("true");
  });

  it("leaves the document and history untouched", () => {
    const { live, secondRow } = tableEditor();
    const original = live.state.doc;
    live.dispatch(live.state.tr.insertText("edited ", secondRow + 3));
    const doc = live.state.doc;
    const selection = live.state.selection;
    expect(undoDepth(live.state)).toBe(1);
    setPageMarks(live, { pushes: [], cuts: [{ at: secondRow, height: 240 }] });
    expect(live.state.doc).toBe(doc);
    expect(live.state.selection).toBe(selection);
    expect(undoDepth(live.state)).toBe(1);
    expect(undo(live.state, live.dispatch)).toBe(true);
    expect(live.state.doc.eq(original)).toBe(true);
    expect(redoDepth(live.state)).toBe(1);

    setPageMarks(live, { pushes: [], cuts: [] });
    expect(live.dom.querySelectorAll("tr")).toHaveLength(2);
    expect(redoDepth(live.state)).toBe(1);
    expect(redo(live.state, live.dispatch)).toBe(true);
    expect(live.state.doc.eq(doc)).toBe(true);
  });

  /**
   * The projection is read off the table as it stands, so the edit alone redraws it. Nothing
   * measured has to be handed back for a header to say what its source row now says.
   */
  it("refreshes a repeated header when its source row changes", () => {
    const { live, secondRow } = tableEditor();
    setPageMarks(live, { pushes: [], cuts: [{ at: secondRow, height: 240 }] });

    let heading = -1;
    live.state.doc.descendants((node, pos) => {
      if (heading < 0 && node.isText && node.text === "Heading A")
        heading = pos;
    });
    if (heading < 0) throw new Error("heading not found");
    live.dispatch(live.state.tr.insertText("Heading C", heading, heading + 9));

    const repeated = live.dom.querySelector(
      `[${editorAttributes.tableRepeatedHeader}]`
    );
    expect(repeated?.textContent).toBe("Heading CHeading B");
  });

  /**
   * One measurement decides a push, a break space and a table continuation together, and the
   * three used to be dispatched one after another. They are one set of marks and one transaction.
   */
  it("a push, a space and a continuation applied together are one transaction", () => {
    const live = mounted(
      docxSchema.nodes.doc.create(null, [
        docxSchema.nodes.paragraph.create({}, [
          docxSchema.text("before"),
          pageBreak(),
          docxSchema.text("after"),
        ]),
        docxSchema.nodes.table.create({ gridCols: [1000, 1000] }, [
          docxSchema.nodes.tableRow.create({ format: { repeatHeader: true } }, [
            cell("H1"),
            cell("H2"),
          ]),
          docxSchema.nodes.tableRow.create({}, [cell("A"), cell("B")]),
        ]),
        docxSchema.nodes.paragraph.create({}, [docxSchema.text("below")]),
      ])
    );
    const dispatched = vi.spyOn(live, "dispatch");
    const tablePos = live.state.doc.child(0).nodeSize;
    const secondRow = tablePos + 1 + live.state.doc.child(1).child(0).nodeSize;

    setPageMarks(live, {
      pushes: [{ pos: tablePos, marginTop: 30, push: 30 }],
      cuts: [
        { at: firstBreak(live), height: 120 },
        { at: secondRow, height: 200 },
      ],
    });

    expect(dispatched).toHaveBeenCalledTimes(1);
    expect(dispatched.mock.calls.every(([tr]) => !tr.docChanged)).toBe(true);
    expect(spaceHeights(live)).toEqual(["120"]);
    expect(
      live.dom
        .querySelector(`[${editorAttributes.tablePageSpace}]`)
        ?.getAttribute(editorAttributes.tablePageSpace)
    ).toBe("200");
  });

  it("applying the same marks again dispatches nothing", () => {
    const { live, secondRow } = tableEditor();
    const marks = { pushes: [], cuts: [{ at: secondRow, height: 240 }] };
    setPageMarks(live, marks);
    const dispatched = vi.spyOn(live, "dispatch");

    setPageMarks(live, marks);
    expect(dispatched).not.toHaveBeenCalled();
  });

  /**
   * A cut names the row its spacer opens before. When that row goes, so does the cut: mapped onto
   * whatever slid into its place, the spacer would open before the wrong row.
   */
  it("a cut whose row was deleted is dropped and the others keep their places", () => {
    const live = mounted(
      docxSchema.nodes.doc.create(null, [
        docxSchema.nodes.table.create({ gridCols: [1000] }, [
          docxSchema.nodes.tableRow.create({ format: { repeatHeader: true } }, [
            cell("Heading"),
          ]),
          docxSchema.nodes.tableRow.create({}, [cell("A")]),
          docxSchema.nodes.tableRow.create({}, [cell("B")]),
          docxSchema.nodes.tableRow.create({}, [cell("C")]),
        ]),
      ])
    );
    const rows: { at: number; size: number }[] = [];
    live.state.doc.child(0).forEach((row, offset) => {
      rows.push({ at: 1 + offset, size: row.nodeSize });
    });
    const rowB = rows[2] ?? { at: 0, size: 0 };
    const rowC = rows[3] ?? { at: 0, size: 0 };

    setPageMarks(live, {
      pushes: [],
      cuts: [
        { at: rowB.at, height: 100 },
        { at: rowC.at, height: 200 },
      ],
    });
    expect(
      Array.from(
        live.dom.querySelectorAll(`[${editorAttributes.tablePageSpace}]`),
        (row) => row.getAttribute(editorAttributes.tablePageSpace)
      )
    ).toEqual(["100", "200"]);

    live.dispatch(live.state.tr.delete(rowB.at, rowB.at + rowB.size));

    const drawn = Array.from(live.dom.querySelectorAll("tr"), (row) =>
      row.hasAttribute(editorAttributes.tablePageSpace)
        ? `space:${row.getAttribute(editorAttributes.tablePageSpace)}`
        : row.textContent
    );
    expect(drawn).toEqual(["Heading", "A", "space:200", "Heading", "C"]);
  });
});

/**
 * No fixture carries a `w:type="page"` (`__fixtures__/README.md`), so this is what says a break
 * read out of a document is the same thing as one a key put in: it comes in as a page break, it is
 * drawn with a space of its own, and it goes back out untouched.
 */
describe("a page break the document arrived with", () => {
  const BODY =
    '<w:p><w:r><w:t xml:space="preserve">before</w:t></w:r>' +
    '<w:r><w:br w:type="page"/></w:r>' +
    '<w:r><w:t xml:space="preserve">after</w:t></w:r></w:p>';

  function imported(): PMNode {
    return importDocx(makeDocx(BODY)).doc;
  }

  it("comes in as a page break", () => {
    const breaks = imported()
      .child(0)
      .children.filter((child) => child.type === docxSchema.nodes.hardBreak);
    expect(breaks.map((child) => isPageBreak(child.attrs.brAttrs))).toEqual([
      true,
    ]);
  });

  it("is given its space before anything is measured", () => {
    expect(spaceHeights(mounted(imported()))).toEqual(["0"]);
  });

  it("goes back out as the `w:br` it came in as", () => {
    // The three runs the paragraph came in as are written back out as one, since nothing tells
    // them apart, so it is the `w:br` itself that has to be identical
    const xml = serializeParagraph(imported().child(0));
    expect(xml.match(/<w:br[^>]*\/>/g)).toEqual(['<w:br w:type="page"/>']);
    expect(xml).toContain('<w:t xml:space="preserve">before</w:t>');
    expect(xml).toContain('<w:t xml:space="preserve">after</w:t>');
  });
});
