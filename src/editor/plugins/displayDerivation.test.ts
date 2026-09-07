// @vitest-environment jsdom
import { undo } from "prosemirror-history";
import {
  DOMSerializer,
  Fragment,
  type Node as PMNode,
} from "prosemirror-model";
import {
  type Command,
  EditorState,
  Plugin,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { afterEach, describe, expect, it } from "vitest";
import {
  fixtureNames,
  makeStyledDocx,
  readFixture,
} from "../../__testing__/docx";
import { select } from "../../__testing__/editing";
import { NO_FORMATTING, styleIdOf } from "../../docx/formatting";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import {
  toCellFormat,
  toParagraphFormat,
  toRunFormat,
} from "../../model/format";
import { docxSchema } from "../../schema";
import { displayOnly } from "../../schema/displayDerivation";
import { editorClassNames } from "../../styles/classNames";
import { addRowAfter } from "../../table";
import {
  setParagraphAlign,
  setParagraphStyle,
} from "../commands/paragraphCommands";
import { createEditorState, editorStateForSession } from "../createEditor";
import {
  documentOf,
  type EditorDocument,
  editorDocument,
  editorDocumentOf,
} from "../editorDocument";
import { insertTable } from "../insertTable";
import { paragraphPPr } from "../paragraphEdits";
import { type DocumentDeriver, displayDerivation } from "./displayDerivation";

/** A default paragraph style that lays down both paragraph and character formatting */
const NORMAL_STYLE =
  '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
  '<w:name w:val="Normal"/>' +
  '<w:pPr><w:jc w:val="center"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>';

const HEADING_STYLE =
  '<w:style w:type="paragraph" w:styleId="Heading1">' +
  '<w:name w:val="heading 1"/>' +
  '<w:rPr><w:i/><w:sz w:val="40"/></w:rPr></w:style>';

const BODY = '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>';

/** The state built out of everything the document told us, the same as `DocxEditor` builds it */
function opened(styles = NORMAL_STYLE): {
  state: EditorState;
  session: SessionStore;
} {
  const { doc, session } = importDocx(makeStyledDocx(BODY, styles));
  return { state: editorStateForSession({ doc, session }), session };
}

function ran(state: EditorState, command: Command): EditorState {
  let next = state;
  expect(
    command(state, (tr) => {
      next = state.apply(tr);
    })
  ).toBe(true);
  return next;
}

function undone(state: EditorState): EditorState {
  let next = state;
  undo(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

/** The paragraph inside the cell at this spot of the first table, and the values it draws with */
function cellParagraph(doc: PMNode, row: number, col: number) {
  const table = doc.child(1);
  const cell = table.child(row).child(col);
  const paragraph = cell.child(0);
  return {
    node: paragraph,
    format: toParagraphFormat(paragraph.attrs.format),
    styleRun: toRunFormat(paragraph.attrs.styleRun),
  };
}

/** Types into the first cell of the table, where inserting one leaves the caret */
function typedInFirstCell(state: EditorState, text: string): EditorState {
  return state.apply(state.tr.insertText(text, state.selection.head));
}

/** The paragraph as it is drawn on screen */
function rendered(paragraph: PMNode): string {
  const host = document.createElement("div");
  host.appendChild(
    DOMSerializer.fromSchema(docxSchema).serializeNode(paragraph)
  );
  return host.innerHTML;
}

/** A paragraph carrying no style values, which is the shape an edit builds one in */
function freshHeading(text = "heading"): PMNode {
  return docxSchema.nodes.paragraph.create(
    { pPr: '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' },
    [docxSchema.text(text)]
  );
}

let view: EditorView | null = null;

afterEach(() => {
  view?.destroy();
  view = null;
});

function mounted(state: EditorState): EditorView {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  view = new EditorView(mount, { state });
  return view;
}

/** The composition events `view.composing` follows, which is all this plugin reads */
function composition(live: EditorView, open: boolean): void {
  live.dom.dispatchEvent(
    new CompositionEvent(open ? "compositionstart" : "compositionend", {
      bubbles: true,
      data: "",
    })
  );
}

/** Puts a fresh paragraph in with the caret inside it, which is where a composition would sit */
function putHeadingUnderTheCaret(live: EditorView): void {
  const at = live.state.doc.child(0).nodeSize;
  const tr = live.state.tr.insert(at, freshHeading());
  live.dispatch(tr.setSelection(TextSelection.create(tr.doc, at + 2)));
}

function headingStyleRun(live: EditorView) {
  return toRunFormat(live.state.doc.child(1).attrs.styleRun);
}

describe("a paragraph an edit built from nothing", () => {
  it("draws text typed into a new table's cell in the default style", () => {
    const inserted = ran(opened().state, insertTable({ rows: 2, columns: 2 }));
    const typed = typedInFirstCell(inserted, "cell");
    const cell = cellParagraph(typed.doc, 0, 0);

    expect(cell.node.textContent).toBe("cell");
    // The text carries no run of its own, so the paragraph is what has to draw it
    expect(cell.node.child(0).marks).toEqual([]);
    expect(cell.styleRun).toEqual({ bold: true, fontSizePt: 11 });
    expect(cell.format).toEqual({ align: "center" });
    expect(rendered(cell.node)).toContain("font-weight: bold");
  });

  it("draws it the same in a cell an added row brought", () => {
    const inserted = ran(opened().state, insertTable({ rows: 2, columns: 2 }));
    const grown = ran(inserted, addRowAfter);
    const typed = typedInFirstCell(grown, "new cell");
    const cell = cellParagraph(typed.doc, 1, 0);

    expect(cell.node.textContent).toBe("new cell");
    expect(cell.styleRun).toEqual({ bold: true, fontSizePt: 11 });
  });

  // The values go in as an appended transaction, which the history takes as part of the same event
  it("is taken back together with the edit that built it, in one undo", () => {
    const state = opened().state;
    const inserted = ran(state, insertTable({ rows: 2, columns: 2 }));
    expect(cellParagraph(inserted.doc, 0, 0).styleRun).toEqual({
      bold: true,
      fontSizePt: 11,
    });

    const back = undone(inserted);
    expect(back.doc.childCount).toBe(1);
    expect(back.doc.child(0).textContent).toBe("body");
  });

  it("is read under the style it names itself", () => {
    const { state } = opened(NORMAL_STYLE + HEADING_STYLE);
    const put = state.apply(
      state.tr.insert(state.doc.child(0).nodeSize, freshHeading())
    );

    expect(toRunFormat(put.doc.child(1).attrs.styleRun)).toEqual({
      italic: true,
      fontSizePt: 20,
    });
  });

  /**
   * An untouched block goes back out as its original XML, and that rests on the node still being
   * the same one, so a paragraph already carrying its values may never be written again
   */
  it("leaves a paragraph that already carries its values as the very same node", () => {
    const inserted = ran(opened().state, insertTable({ rows: 2, columns: 2 }));
    const body = inserted.doc.child(0);
    const typed = typedInFirstCell(inserted, "cell");
    const untouchedCell = cellParagraph(typed.doc, 1, 1).node;

    // The paragraph the document came with is untouched by the cells being filled in
    expect(typed.doc.child(0)).toBe(body);

    const again = typed.apply(
      typed.tr.insertText("more", typed.selection.head)
    );
    expect(again.doc.child(0)).toBe(body);
    expect(cellParagraph(again.doc, 1, 1).node).toBe(untouchedCell);
  });
});

/**
 * Rewriting a node redraws it, and a redraw under an open composition is what the browser answers
 * by taking the composition down. So the paragraph an IME is composing in waits.
 */
describe("a paragraph an IME is composing in", () => {
  it("is written to as usual while nothing is being composed", () => {
    const live = mounted(opened(NORMAL_STYLE + HEADING_STYLE).state);
    putHeadingUnderTheCaret(live);

    expect(headingStyleRun(live)).toEqual({ italic: true, fontSizePt: 20 });
  });

  it("is left as it stands until the composition is over", () => {
    const live = mounted(opened(NORMAL_STYLE + HEADING_STYLE).state);
    composition(live, true);
    expect(live.composing).toBe(true);

    putHeadingUnderTheCaret(live);
    expect(headingStyleRun(live)).toBeNull();

    // Every further edit under the same composition leaves it alone as well
    live.dispatch(live.state.tr.insertText("more", live.state.selection.head));
    expect(headingStyleRun(live)).toBeNull();

    // And the next edit after the composition is the one that reads the styles into it
    composition(live, false);
    expect(live.composing).toBe(false);
    live.dispatch(live.state.tr.insertText("!", 1));

    expect(headingStyleRun(live)).toEqual({ italic: true, fontSizePt: 20 });
  });

  it("does not hold up the other paragraphs of the document", () => {
    const live = mounted(opened(NORMAL_STYLE + HEADING_STYLE).state);
    composition(live, true);
    putHeadingUnderTheCaret(live);

    // A second fresh paragraph, this one nowhere near the composition
    const tr = live.state.tr.insert(
      live.state.doc.content.size,
      freshHeading()
    );
    live.dispatch(tr);

    expect(headingStyleRun(live)).toBeNull();
    expect(toRunFormat(live.state.doc.child(2).attrs.styleRun)).toEqual({
      italic: true,
      fontSizePt: 20,
    });
  });
});

/** Every paragraph of the document in the order they stand, table cells included */
function paragraphsOf(doc: PMNode): PMNode[] {
  const found: PMNode[] = [];
  doc.descendants((node) => {
    if (node.type === docxSchema.nodes.paragraph) found.push(node);
    return node.type !== docxSchema.nodes.paragraph;
  });
  return found;
}

/** Records every re-derivation the plugin appends, which is what the gate tests count */
function watching(): { plugin: Plugin; appended: Transaction[] } {
  const appended: Transaction[] = [];
  return {
    appended,
    plugin: new Plugin({
      filterTransaction(tr) {
        if (tr.getMeta(displayOnly) === true) appended.push(tr);
        return true;
      },
    }),
  };
}

const lockedPr =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

/** `w:sz` counts eighths of a point, so 8 is 1pt and 4 is 0.5pt */
const borderXml = (side: string, eighths: number) =>
  `<w:${side} w:val="single" w:sz="${eighths}" w:color="000000"/>`;

/** A table with a 1pt line around the outside and 0.5pt lines between its cells */
const TBL_PR =
  "<w:tblPr><w:tblBorders>" +
  ["top", "left", "bottom", "right"]
    .map((side) => borderXml(side, 8))
    .join("") +
  borderXml("insideH", 4) +
  borderXml("insideV", 4) +
  "</w:tblBorders></w:tblPr>";

const GRID =
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>';

/** The body and a two-column table whose left cell stands inside a control that shuts it */
const BODY_WITH_LOCKED_CELL =
  BODY +
  `<w:tbl>${GRID}<w:tr><w:sdt>${lockedPr}<w:sdtContent>${cellXml("Locked")}</w:sdtContent></w:sdt>` +
  `${cellXml("Open")}</w:tr></w:tbl>`;

/** The body and a lined two by two table */
const BODY_WITH_LINED_TABLE =
  BODY +
  `<w:tbl>${TBL_PR}${GRID}` +
  `<w:tr>${cellXml("a1")}${cellXml("b1")}</w:tr>` +
  `<w:tr>${cellXml("a2")}${cellXml("b2")}</w:tr></w:tbl>`;

describe("what the state is built over", () => {
  it.each(fixtureNames)(
    "%s: opened under its own session, every paragraph is the very node the import built",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const state = editorStateForSession({ doc, session });
      const imported = paragraphsOf(doc);
      const held = paragraphsOf(state.doc);

      expect(held).toHaveLength(imported.length);
      held.forEach((paragraph, at) => {
        expect(paragraph).toBe(imported[at]);
      });
    }
  );

  it("opened under another snapshot, every paragraph is worked out against it with its properties untouched", () => {
    const { doc, session } = importDocx(makeStyledDocx(BODY, NORMAL_STYLE));
    expect(toRunFormat(doc.child(0).attrs.styleRun)).toEqual({
      bold: true,
      fontSizePt: 11,
    });

    const under = createEditorState(doc, {
      document: { ...editorDocumentOf(session), formatting: NO_FORMATTING },
    });
    const body = under.doc.child(0);
    expect(toRunFormat(body.attrs.styleRun)).toBeNull();
    expect(toParagraphFormat(body.attrs.format)).toBeNull();
    expect(body.attrs.pPr).toBe(doc.child(0).attrs.pPr);
    expect(body.attrs.srcId).toBe(doc.child(0).attrs.srcId);
  });
});

/**
 * A style edit will one day replace the snapshot through the document node's attrs
 * (`editor/editorDocument`). What stands in for it here is a state opened under no formatting at
 * all, whose session answers with the document's own once the snapshot is derived again.
 */
describe("the snapshot being replaced under the document", () => {
  function underNothing(body = BODY_WITH_LOCKED_CELL): {
    doc: PMNode;
    state: EditorState;
    appended: Transaction[];
  } {
    const { doc, session } = importDocx(makeStyledDocx(body, NORMAL_STYLE));
    const document: EditorDocument = {
      ...editorDocumentOf(session),
      formatting: NO_FORMATTING,
    };
    const { plugin, appended } = watching();
    return {
      doc,
      state: createEditorState(doc, { document, consumerPlugins: [plugin] }),
      appended,
    };
  }

  /** Moves the document node's attrs, which is what derives the snapshot from the session again */
  const restyle = (tr: Transaction) => tr.setDocAttribute("sectPr", null);

  function restyled(state: EditorState): EditorState {
    const next = state.apply(restyle(state.tr));
    expect(documentOf(next)).not.toBe(documentOf(state));
    return next;
  }

  const NORMAL_RUN = { bold: true, fontSizePt: 11 };

  it("works every paragraph out against the new snapshot without touching its properties", () => {
    const { doc, state } = underNothing();
    for (const paragraph of paragraphsOf(state.doc)) {
      expect(paragraph.attrs.styleRun).toBeNull();
    }

    const held = paragraphsOf(restyled(state).doc);
    expect(held).toHaveLength(3);
    held.forEach((paragraph, at) => {
      expect(toRunFormat(paragraph.attrs.styleRun)).toEqual(NORMAL_RUN);
      expect(toParagraphFormat(paragraph.attrs.format)).toEqual({
        align: "center",
      });
      expect(paragraph.attrs.pPr).toBe(paragraphsOf(doc)[at]?.attrs.pPr);
    });
  });

  it("re-derives the paragraph inside a locked cell as well", () => {
    const { state } = underNothing();
    const locked = paragraphsOf(restyled(state).doc).find(
      (paragraph) => paragraph.textContent === "Locked"
    );

    expect(toRunFormat(locked?.attrs.styleRun)).toEqual(NORMAL_RUN);
  });

  it("goes to the history not at all, where the re-derivation after an edit goes with the edit", () => {
    const { state, appended } = underNothing();
    const next = restyled(state);
    expect(appended).toHaveLength(1);
    expect(appended[0]?.getMeta("addToHistory")).toBe(false);

    // A new cell's paragraph is read into the style the snapshot now lays down
    appended.length = 0;
    ran(next, insertTable({ rows: 1, columns: 1 }));
    expect(appended).toHaveLength(1);
    expect(appended[0]?.getMeta("addToHistory")).toBeUndefined();
  });

  it("re-derives the paragraph being composed in as well, since nothing may stay under the old snapshot", () => {
    const live = mounted(underNothing(BODY).state);
    composition(live, true);
    live.dispatch(restyle(live.state.tr));

    expect(toRunFormat(live.state.doc.child(0).attrs.styleRun)).toEqual(
      NORMAL_RUN
    );
  });
});

describe("what a node maps back to", () => {
  /** A deriver that writes nothing and records what each paragraph was handed as the node it was */
  function recording(): {
    deriver: DocumentDeriver;
    handed: Map<string, PMNode | null>;
  } {
    const handed = new Map<string, PMNode | null>();
    return {
      handed,
      deriver: {
        name: "recording",
        nodeTypes: ["paragraph"],
        derive(node, _pos, _doc, _context, previous) {
          handed.set(node.textContent, previous);
          return [];
        },
      },
    };
  }

  const twoParagraphs = () =>
    docxSchema.nodes.doc.create(null, [
      docxSchema.nodes.paragraph.create({ pPr: "<w:pPr/>" }, [
        docxSchema.text("first"),
      ]),
      docxSchema.nodes.paragraph.create(null, [docxSchema.text("second")]),
    ]);

  function watched(deriver: DocumentDeriver): EditorState {
    return EditorState.create({
      doc: twoParagraphs(),
      plugins: [displayDerivation([deriver])],
    });
  }

  it("is the node itself once it was rewritten where it stands", () => {
    const { deriver, handed } = recording();
    const state = watched(deriver);
    const [first, second] = [state.doc.child(0), state.doc.child(1)];
    state.apply(
      state.tr.setNodeMarkup(0, null, {
        ...first.attrs,
        pPr: '<w:pPr><w:jc w:val="center"/></w:pPr>',
      })
    );

    expect(handed.get("first")).toBe(first);
    expect(handed.get("second")).toBe(second);
  });

  it("is the node that stood there once something was put in ahead of it, and none for what was put in", () => {
    const { deriver, handed } = recording();
    const state = watched(deriver);
    const first = state.doc.child(0);
    state.apply(
      state.tr.insert(
        0,
        docxSchema.nodes.paragraph.create(null, [docxSchema.text("new")])
      )
    );

    expect(handed.get("new")).toBeNull();
    expect(handed.get("first")).toBe(first);
  });

  it("is the paragraph for the half of a split that carries on, and none for the half the split made", () => {
    const { deriver, handed } = recording();
    const state = watched(deriver);
    const first = state.doc.child(0);
    state.apply(state.tr.split(3));

    expect(handed.get("fi")).toBe(first);
    expect(handed.get("rst")).toBeNull();
  });

  it("is none for every node once the snapshot is replaced", () => {
    const { deriver, handed } = recording();
    const { doc, session } = importDocx(makeStyledDocx(BODY, NORMAL_STYLE));
    const state = EditorState.create({
      doc,
      plugins: [
        editorDocument(editorDocumentOf(session)),
        displayDerivation([deriver]),
      ],
    });
    state.apply(state.tr.setDocAttribute("sectPr", null));

    expect(handed.get("body")).toBeNull();
  });

  /** The table copied cell by cell, each cell carrying no lines of its own */
  function withoutCellLines(table: PMNode): PMNode {
    return table.copy(
      Fragment.from(
        table.children.map((row) =>
          row.copy(
            Fragment.from(
              row.children.map((cell) =>
                cell.type.create({ ...cell.attrs, format: null }, cell.content)
              )
            )
          )
        )
      )
    );
  }

  const cellLines = (table: PMNode): (string | undefined)[] =>
    table.children.flatMap((row) =>
      row.children.map((cell) => toCellFormat(cell.attrs.format)?.borderTop)
    );

  /**
   * Paired by the order they stand in, the new table would be mistaken for the one it was
   * copied from, which reads the same in every input its lines depend on, and would keep the
   * lines it arrived without.
   */
  it("derives a table put in ahead of another of the same shape rather than mistaking it for that one", () => {
    const { doc, session } = importDocx(
      makeStyledDocx(BODY_WITH_LINED_TABLE, NORMAL_STYLE)
    );
    const state = editorStateForSession({ doc, session });
    const lined = state.doc.child(1);
    expect(cellLines(lined)).toEqual([
      "1pt solid #000000",
      "1pt solid #000000",
      "0.5pt solid #000000",
      "0.5pt solid #000000",
    ]);

    const inserted = state.apply(
      state.tr.insert(state.doc.child(0).nodeSize, withoutCellLines(lined))
    );

    expect(cellLines(inserted.doc.child(1))).toEqual(cellLines(lined));
    expect(inserted.doc.child(2)).toBe(lined);
  });
});

/**
 * A command that rewrites a paragraph's properties writes the display values with them, along
 * the same path the deriver takes, so the re-derivation that follows finds nothing to write. This
 * is what keeps the registry from changing what any of them does.
 */
describe("a paragraph a command already wrote the values of", () => {
  function openedWatching(): { state: EditorState; appended: Transaction[] } {
    const { doc, session } = importDocx(
      makeStyledDocx(BODY, NORMAL_STYLE + HEADING_STYLE)
    );
    const { plugin, appended } = watching();
    return {
      state: editorStateForSession(
        { doc, session },
        { consumerPlugins: [plugin] }
      ),
      appended,
    };
  }

  it("is left as the style command wrote it", () => {
    const { state, appended } = openedWatching();
    const styled = ran(select(state, 2), setParagraphStyle("Heading1"));

    expect(toRunFormat(styled.doc.child(0).attrs.styleRun)).toEqual({
      italic: true,
      fontSizePt: 20,
    });
    expect(appended).toEqual([]);
  });

  it("is left as the paragraph edit wrote it", () => {
    const { state, appended } = openedWatching();
    const aligned = ran(select(state, 2), setParagraphAlign("right"));

    expect(toParagraphFormat(aligned.doc.child(0).attrs.format)).toEqual({
      align: "right",
    });
    expect(appended).toEqual([]);
  });

  it("is left as the paste wrote it", () => {
    const { state, appended } = openedWatching();
    const live = mounted(select(state, 2));
    const event = new Event("paste", { bubbles: true, cancelable: true });
    Object.defineProperty(event, "clipboardData", {
      value: {
        getData: (type: string) =>
          type === "text/html"
            ? `<p class="${editorClassNames.paragraph}" ` +
              "data-ppr='<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>'>pasted</p>"
            : "",
      },
    });
    live.dom.dispatchEvent(event);

    expect(live.state.doc.textContent).toContain("pasted");
    const pasted = paragraphsOf(live.state.doc).find((paragraph) =>
      paragraph.textContent.includes("pasted")
    );
    expect(styleIdOf(paragraphPPr(pasted ?? live.state.doc.child(0)))).toBe(
      "Heading1"
    );
    expect(toRunFormat(pasted?.attrs.styleRun)).toEqual({
      italic: true,
      fontSizePt: 20,
    });
    expect(appended).toEqual([]);
  });
});
