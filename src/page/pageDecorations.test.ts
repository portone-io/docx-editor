// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import { makeDocx } from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { serializeParagraph } from "../docx/serializeParagraph";
import { createEditorState } from "../editor/createEditor";
import { docxSchema, isPageBreak } from "../schema";
import { editorAttributes } from "../styles/classNames";
import {
  setPageBreakSpaces,
  setPagePushes,
  setTableContinuations,
} from "./pageDecorations";

let view: EditorView | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
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
      docxSchema.nodes.paragraph.create({ format: { spaceBeforePt: 12 } }, [
        docxSchema.text("first paragraph"),
      ]),
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

describe("setPagePushes", () => {
  it("only the pushed block gets a wider gap", () => {
    const live = editor();
    setPagePushes(live, [
      { pos: secondBlock(live), marginTop: 300, push: 300 },
    ]);

    const [first, second] = paragraphs(live);
    expect(pushed(first)).toBe("");
    expect(pushed(second)).toBe("300px");
    expect(second?.getAttribute("data-page-push")).toBe("300");
  });

  it("the gap the paragraph already had of its own remains after the push is taken away", () => {
    const live = editor();
    setPagePushes(live, [{ pos: 0, marginTop: 316, push: 300 }]);
    expect(pushed(paragraphs(live)[0])).toBe("316px");
    expect(own(paragraphs(live)[0])).toBe("12pt");

    setPagePushes(live, []);
    expect(pushed(paragraphs(live)[0])).toBe("");
    expect(own(paragraphs(live)[0])).toBe("12pt");
    expect(paragraphs(live)[0]?.hasAttribute("data-page-push")).toBe(false);
  });

  it("does not touch the document when the same value is applied again", () => {
    const live = editor();
    setPagePushes(live, [{ pos: 0, marginTop: 16, push: 16 }]);
    const before = live.state;
    setPagePushes(live, [{ pos: 0, marginTop: 16, push: 16 }]);
    expect(live.state).toBe(before);
  });

  it("a push leaves nothing behind in the document or in the edit history", () => {
    const live = editor();
    const doc = live.state.doc;
    setPagePushes(live, [{ pos: 0, marginTop: 40, push: 24 }]);
    expect(live.state.doc).toBe(doc);
  });
});

/** One editor with a single paragraph broken twice, so the breaks have ordinals that can shift */
function brokenEditor(): EditorView {
  const pageBreak = () =>
    docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' });
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

/** The height standing on each break's space, in document order */
function spaceHeights(live: EditorView): string[] {
  return Array.from(
    live.dom.querySelectorAll(`[${editorAttributes.pageBreakSpace}]`),
    (element) => element.getAttribute(editorAttributes.pageBreakSpace) ?? ""
  );
}

/** Where the first `br` of the document stands */
function firstBreak(live: EditorView): number {
  let found = -1;
  live.state.doc.descendants((node, pos) => {
    if (found < 0 && node.type === docxSchema.nodes.hardBreak) found = pos;
  });
  if (found < 0) throw new Error("the document holds no break");
  return found;
}

/** The two heights the layout works out for the two breaks */
const TWO_SPACES = [
  { pos: 0, index: 0, height: 111 },
  { pos: 0, index: 1, height: 222 },
];

describe("setPageBreakSpaces", () => {
  it("opens each break the height it was given", () => {
    const live = brokenEditor();
    setPageBreakSpaces(live, TWO_SPACES);
    expect(spaceHeights(live)).toEqual(["111", "222"]);
  });

  it("every break is left with an empty space when they are taken away", () => {
    const live = brokenEditor();
    setPageBreakSpaces(live, TWO_SPACES);

    setPageBreakSpaces(live, []);
    expect(spaceHeights(live)).toEqual(["0", "0"]);
  });

  it("does not touch the document when the same values are applied again", () => {
    const live = brokenEditor();
    setPageBreakSpaces(live, TWO_SPACES);
    const before = live.state;
    setPageBreakSpaces(live, TWO_SPACES);
    expect(live.state).toBe(before);
  });

  it("a space leaves nothing behind in the document or in the edit history", () => {
    const live = brokenEditor();
    const doc = live.state.doc;
    setPageBreakSpaces(live, TWO_SPACES);
    expect(live.state.doc).toBe(doc);
  });

  /**
   * A space is measured for one break and would be a page-sized height on any other, so it has to
   * follow the very `br` it was worked out for. Keyed by the ordinal the layout counted, the
   * survivor of a deletion took the height of the break that went away.
   */
  it("a space stays with its own break when an earlier break is deleted", () => {
    const live = brokenEditor();
    setPageBreakSpaces(live, TWO_SPACES);

    const first = firstBreak(live);
    live.dispatch(live.state.tr.delete(first, first + 1));

    expect(spaceHeights(live)).toEqual(["222"]);
  });

  it("a break put in before another one starts out with no space of its own", () => {
    const live = brokenEditor();
    setPageBreakSpaces(live, TWO_SPACES);

    live.dispatch(
      live.state.tr.insert(
        firstBreak(live),
        docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' })
      )
    );

    expect(spaceHeights(live)).toEqual(["0", "111", "222"]);
  });
});

function tableEditor(): {
  live: EditorView;
  secondRow: number;
  headerRow: number;
} {
  const paragraph = (text: string) =>
    docxSchema.nodes.paragraph.create({}, [docxSchema.text(text)]);
  const cell = (text: string) =>
    docxSchema.nodes.tableCell.create({}, [paragraph(text)]);
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

describe("setTableContinuations", () => {
  it("inserts a non-editable spacer and header projection before the continued row", () => {
    const { live, headerRow, secondRow } = tableEditor();
    setTableContinuations(live, [
      {
        pos: secondRow,
        height: 240,
        headerRows: [headerRow],
        headerSignature: "heading-a",
        columns: 2,
      },
    ]);

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
    const { live, headerRow, secondRow } = tableEditor();
    const doc = live.state.doc;
    setTableContinuations(live, [
      {
        pos: secondRow,
        height: 240,
        headerRows: [headerRow],
        headerSignature: "heading-a",
        columns: 2,
      },
    ]);
    expect(live.state.doc).toBe(doc);

    setTableContinuations(live, []);
    expect(live.dom.querySelectorAll("tr")).toHaveLength(2);
  });

  it("refreshes a repeated header when its source row changes", () => {
    const { live, headerRow, secondRow } = tableEditor();
    const continuation = {
      pos: secondRow,
      height: 240,
      headerRows: [headerRow],
      headerSignature: "heading-a",
      columns: 2,
    };
    setTableContinuations(live, [continuation]);

    let heading = -1;
    live.state.doc.descendants((node, pos) => {
      if (heading < 0 && node.isText && node.text === "Heading A")
        heading = pos;
    });
    if (heading < 0) throw new Error("heading not found");
    live.dispatch(live.state.tr.insertText("Heading C", heading, heading + 9));
    setTableContinuations(live, [
      { ...continuation, headerSignature: "heading-c" },
    ]);

    const repeated = live.dom.querySelector(
      `[${editorAttributes.tableRepeatedHeader}]`
    );
    expect(repeated?.textContent).toBe("Heading CHeading B");
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
