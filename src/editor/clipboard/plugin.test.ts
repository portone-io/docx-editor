// @vitest-environment jsdom

import type { Node as PMNode } from "prosemirror-model";
import { AllSelection, Plugin, TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";
import { beforeEach, describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeNumberedDocx,
  makeStyledDocx,
  TINY_PNG_DATA_URL,
} from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { toRunFormat } from "../../model/format";
import { emuToPx } from "../../ooxml/image";
import { styleIdOf } from "../../ooxml/props";
import { docxSchema } from "../../schema";
import { editorClassNames } from "../../styles/classNames";
import {
  createEditorState,
  createEditorView,
  editorStateForSession,
} from "../createEditor";
import { defineClipboardEvent } from "./__testing__/clipboardEvent";

const HEADING_STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
  '<w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="DocumentHeading1">' +
  '<w:name w:val="heading 1"/><w:pPr><w:spacing w:before="240"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>';

const PARAGRAPH =
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>';

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

const TABLE =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
  `<w:tr>${cellXml("A1")}${cellXml("B1")}</w:tr>` +
  `<w:tr>${cellXml("A2")}${cellXml("B2")}</w:tr>` +
  "</w:tbl>";

function openEditor(
  bytes: Uint8Array,
  plugins: readonly Plugin[] = []
): { view: EditorView; session: SessionStore } {
  const { doc, session } = importDocx(bytes);
  const view = createEditorView({
    mount: document.createElement("div"),
    state: editorStateForSession(
      { doc, session },
      { consumerPlugins: plugins }
    ),
    onStateChange: () => {},
  });
  return { view, session };
}

/** A paste the way the browser delivers one, since jsdom has no clipboard of its own */
function paste(view: EditorView, data: Record<string, string>): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => data[type] ?? "" },
  });
  view.dom.dispatchEvent(event);
}

/** Holding Shift while pasting, which is the keypress ProseMirror reads the plain paste off */
function pressShift(view: EditorView): void {
  view.dom.dispatchEvent(
    new KeyboardEvent("keydown", { key: "V", shiftKey: true, bubbles: true })
  );
}

function cellTexts(doc: PMNode): string[] {
  const texts: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "tableCell") texts.push(node.textContent);
    return true;
  });
  return texts;
}

function selectAllCells(view: EditorView): void {
  const cells: number[] = [];
  view.state.doc.descendants((node, pos) => {
    if (node.type.name === "tableCell") cells.push(pos);
    return true;
  });
  const first = cells[0];
  const last = cells[cells.length - 1];
  if (first === undefined || last === undefined) {
    throw new Error("the document holds no table");
  }
  view.dispatch(
    view.state.tr.setSelection(
      CellSelection.create(view.state.doc, first, last)
    )
  );
}

/** The whole of the one paragraph the documents here hold, which is what a paste replaces */
function selectSource(view: EditorView): void {
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
  );
}

function openStyledEditor(): { view: EditorView; session: SessionStore } {
  const opened = openEditor(makeStyledDocx(PARAGRAPH, HEADING_STYLES));
  selectSource(opened.view);
  return opened;
}

/** A document holding one paragraph with everything the schema draws privately */
function openLoadedEditor(): EditorView {
  const paragraph = docxSchema.nodes.paragraph.create(
    {
      pPr: '<w:pPr><w:pStyle w:val="DocumentHeading1"/></w:pPr>',
      pAttrs: 'w:rsidR="00AB12"',
    },
    [
      docxSchema.nodes.commentStart.create({ id: "0" }),
      docxSchema.text("commented", [
        docxSchema.marks.run.create({
          rPr: "<w:rPr><w:b/></w:rPr>",
          format: { bold: true },
        }),
      ]),
      docxSchema.nodes.commentEnd.create({ id: "0" }),
      docxSchema.nodes.commentReference.create({
        id: "0",
        author: "Jane Doe",
        authorId: "jane-identity",
        text: "what the comment says",
        commentXml: "<w:comment>anything</w:comment>",
        replies: [{ id: "1", author: "Bob", text: "what the reply says" }],
      }),
      docxSchema.nodes.noteReference.create({
        id: "2",
        kind: "footnote",
        label: "1",
        text: "what the footnote says",
      }),
    ]
  );
  const view = createEditorView({
    mount: document.createElement("div"),
    state: createEditorState(docxSchema.nodes.doc.create(null, [paragraph])),
    onStateChange: () => {},
  });
  view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
  return view;
}

function copiedHtml(view: EditorView): string {
  return view.serializeForClipboard(view.state.selection.content()).dom
    .innerHTML;
}

/** Everything a document leaves on the clipboard as text */
function copiedText(doc: PMNode): string {
  const view = createEditorView({
    mount: document.createElement("div"),
    state: createEditorState(doc),
    onStateChange: () => {},
  });
  view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
  const { text } = view.serializeForClipboard(view.state.selection.content());
  view.destroy();
  return text;
}

beforeEach(() => {
  defineClipboardEvent();
});

describe("the clipboard plugin", () => {
  it("routes view.pasteHTML through the same reader as a paste event", () => {
    const html = "<h1>Heading</h1>";
    const byEvent = openEditor(makeStyledDocx(PARAGRAPH, HEADING_STYLES));
    paste(byEvent.view, { "text/html": html, "text/plain": "Heading" });
    const byCall = openEditor(makeStyledDocx(PARAGRAPH, HEADING_STYLES));
    byCall.view.pasteHTML(html);

    expect(documentXmlOf(byCall.view.state.doc, byCall.session)).toBe(
      documentXmlOf(byEvent.view.state.doc, byEvent.session)
    );
    expect(documentXmlOf(byCall.view.state.doc, byCall.session)).toContain(
      '<w:pStyle w:val="DocumentHeading1"/>'
    );
    byEvent.view.destroy();
    byCall.view.destroy();
  });

  it("never runs a schema parseDOM rule on clipboard DOM", () => {
    const { view } = openEditor(makeDocx(PARAGRAPH));
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
    );
    paste(view, {
      "text/plain": "smuggled",
      "text/html":
        `<p class="${editorClassNames.paragraph}" data-pm-slice="0 0 []" ` +
        'data-ppr="&lt;w:pPr&gt;&lt;w:jc w:val=&quot;center&quot;/&gt;&lt;/w:pPr&gt;" ' +
        'data-pattrs="w:rsidR=&quot;00AB12&quot;">smuggled</p>',
    });

    const paragraph = view.state.doc.firstChild;
    expect(paragraph?.textContent).toBe("smuggled");
    // Both would have come over as they stand had the schema's own rule read this markup
    expect(paragraph?.attrs.pPr).toBeNull();
    expect(paragraph?.attrs.pAttrs).toBeNull();
    view.destroy();
  });

  it("closes a foreign block paste so a heading lands as its own paragraph", () => {
    const { view } = openEditor(makeStyledDocx(PARAGRAPH, HEADING_STYLES));
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 4))
    );
    paste(view, { "text/plain": "Heading", "text/html": "<h1>Heading</h1>" });

    expect(
      view.state.doc.children.map((paragraph) => paragraph.textContent)
    ).toEqual(["sou", "Heading", "rce"]);
    view.destroy();
  });

  it("pastes a paragraph over a cell selection into every selected cell", () => {
    const { view } = openEditor(makeDocx(TABLE));
    selectAllCells(view);
    paste(view, { "text/plain": "x", "text/html": "<p>x</p>" });

    expect(cellTexts(view.state.doc)).toEqual(["x", "x", "x", "x"]);
    view.destroy();
  });

  it.each([false, true])(
    "uses the text fallback for unreadable HTML (cell selection: %s)",
    (cells) => {
      const { view } = openEditor(makeDocx(cells ? TABLE : PARAGRAPH));
      if (cells) selectAllCells(view);
      else selectSource(view);
      paste(view, {
        "text/html": "<!-- clipboard producer metadata -->",
        "text/plain": "fallback",
      });
      if (cells)
        expect(cellTexts(view.state.doc)).toEqual(Array(4).fill("fallback"));
      else expect(view.state.doc.textContent).toBe("fallback");
      view.destroy();
    }
  );

  it.each([false, true])(
    "keeps the selection when a paste has no usable content (cell selection: %s)",
    (cells) => {
      const { view } = openEditor(makeDocx(cells ? TABLE : PARAGRAPH));
      if (cells) selectAllCells(view);
      else selectSource(view);
      const before = view.state;
      const cases: Record<string, string>[] = [
        { "text/plain": "\u0001" },
        { "text/html": "<!-- no content -->" },
      ];
      for (const data of cases) {
        paste(view, data);
        expect(view.state.doc.eq(before.doc)).toBe(true);
        expect(view.state.selection.eq(before.selection)).toBe(true);
      }
      view.destroy();
    }
  );

  it("keeps a Shift paste as plain text with w:br line breaks", () => {
    const { view, session } = openEditor(makeDocx(PARAGRAPH));
    view.dispatch(
      view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
    );
    pressShift(view);
    paste(view, {
      "text/plain": "a\nb",
      "text/html": '<p><span style="font-weight: 700">a<br>b</span></p>',
    });

    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.textContent).toBe("ab");
    expect(documentXmlOf(view.state.doc, session)).toContain("<w:br/>");
    let bold = false;
    view.state.doc.descendants((node) => {
      for (const mark of node.marks) {
        if (toRunFormat(mark.attrs.format)?.bold === true) bold = true;
      }
      return true;
    });
    expect(bold).toBe(false);
    view.destroy();
  });

  it("dispatches the paste with the uiEvent meta", () => {
    const events: unknown[] = [];
    const watcher = new Plugin({
      appendTransaction(transactions) {
        for (const tr of transactions) {
          if (tr.docChanged) events.push(tr.getMeta("uiEvent"));
        }
        return null;
      },
    });
    const { view } = openEditor(makeDocx(PARAGRAPH), [watcher]);
    paste(view, { "text/plain": "pasted", "text/html": "<p>pasted</p>" });

    // Nonempty content reaches ProseMirror's insertion, even though the plugin handles empty
    // readings and their text fallback before tableEditing can consume an empty slice.
    expect(events).toEqual(["paste"]);
    view.destroy();
  });
});

describe("copying out of the editor", () => {
  it("copied HTML carries no data- attribute other than data-style", () => {
    const view = openLoadedEditor();
    const html = copiedHtml(view);

    // `data-pm-slice` is prosemirror-view's own, added after the document is drawn
    expect(html.match(/\sdata-(?!style=|pm-slice=)[\w-]+/g)).toBeNull();
    expect(html).not.toContain("<w:");
    view.destroy();
  });

  it("copied HTML carries no comment author or note body", () => {
    const view = openLoadedEditor();
    const html = copiedHtml(view);

    for (const secret of [
      "Jane Doe",
      "jane-identity",
      "what the comment says",
      "what the reply says",
      "what the footnote says",
    ]) {
      expect(html).not.toContain(secret);
    }
    view.destroy();
  });

  it("says an image's size in the pixels it was drawn at", () => {
    const extent = { cx: 1905000, cy: 952500 };
    const paragraph = docxSchema.nodes.paragraph.create(null, [
      docxSchema.nodes.image.create({ src: TINY_PNG_DATA_URL, extent }),
    ]);
    const view = createEditorView({
      mount: document.createElement("div"),
      state: createEditorState(docxSchema.nodes.doc.create(null, [paragraph])),
      onStateChange: () => {},
    });
    view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
    const html = copiedHtml(view);

    // The measure the document keeps does not travel; the pixels the browser was given do, and
    // `editor/plugins/imagePaste` reads a size back out of them
    expect(html).not.toContain("data-extent");
    expect(html).toContain(`width="${Math.round(emuToPx(extent.cx))}"`);
    expect(html).toContain(`height="${Math.round(emuToPx(extent.cy))}"`);
    view.destroy();
  });

  it("copied text keeps a run of tabs as the run it was", () => {
    const tabbed = (count: number) =>
      docxSchema.nodes.doc.create(null, [
        docxSchema.nodes.paragraph.create(null, [
          docxSchema.text("A"),
          docxSchema.text("\t".repeat(count), [docxSchema.marks.tab.create()]),
          docxSchema.text("B"),
        ]),
      ]);

    expect(copiedText(tabbed(1))).toBe("A\tB");
    expect(copiedText(tabbed(2))).toBe("A\t\tB");
    expect(copiedText(tabbed(3))).toBe("A\t\t\tB");
  });

  it("copied text leaves the number off a note that draws its own mark", () => {
    const noted = (customMarkFollows: boolean) =>
      docxSchema.nodes.doc.create(null, [
        docxSchema.nodes.paragraph.create(null, [
          docxSchema.text("Text"),
          docxSchema.nodes.noteReference.create({
            id: "2",
            kind: "footnote",
            label: "1",
            text: "the body",
            customMarkFollows,
          }),
          ...(customMarkFollows ? [docxSchema.text("*")] : []),
        ]),
      ]);

    expect(copiedText(noted(true))).toBe("Text*");
    expect(copiedText(noted(false))).toBe("Text1");
  });

  it("copied text keeps tabs, breaks and cell boundaries", () => {
    const cell = (text: string) =>
      docxSchema.nodes.tableCell.create(null, [
        docxSchema.nodes.paragraph.create(null, [docxSchema.text(text)]),
      ]);
    const row = (left: string, right: string) =>
      docxSchema.nodes.tableRow.create(null, [cell(left), cell(right)]);
    const tabbed = docxSchema.nodes.paragraph.create(null, [
      docxSchema.text("before"),
      docxSchema.text("\t", [docxSchema.marks.tab.create({})]),
      docxSchema.text("after"),
      docxSchema.nodes.hardBreak.create({ brAttrs: null }),
      docxSchema.text("next line"),
      docxSchema.nodes.hardBreak.create({ brAttrs: 'w:type="page"' }),
      docxSchema.text("next page"),
    ]);
    const view = createEditorView({
      mount: document.createElement("div"),
      state: createEditorState(
        docxSchema.nodes.doc.create(null, [
          tabbed,
          docxSchema.nodes.table.create(null, [
            row("one", "two"),
            row("three", "four"),
          ]),
        ])
      ),
      onStateChange: () => {},
    });
    view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
    const { text } = view.serializeForClipboard(view.state.selection.content());

    expect(text).toContain("before\tafter");
    expect(text).toContain("after\nnext line");
    expect(text).toContain("next line\fnext page");
    expect(text).toContain("one\ttwo");
    expect(text).toContain("one\ttwo\nthree\tfour");
    view.destroy();
  });

  it("copied text says nothing for a comment marker and speaks for an image", () => {
    const paragraph = docxSchema.nodes.paragraph.create(null, [
      docxSchema.nodes.commentStart.create({ id: "0" }),
      docxSchema.text("body"),
      docxSchema.nodes.commentEnd.create({ id: "0" }),
      docxSchema.nodes.commentReference.create({
        id: "0",
        author: "Jane Doe",
        text: "what the comment says",
      }),
      docxSchema.nodes.image.create({
        src: TINY_PNG_DATA_URL,
        alt: "a picture of a cat",
      }),
      docxSchema.nodes.noteReference.create({
        id: "2",
        kind: "footnote",
        label: "7",
        text: "what the footnote says",
      }),
    ]);
    const view = createEditorView({
      mount: document.createElement("div"),
      state: createEditorState(docxSchema.nodes.doc.create(null, [paragraph])),
      onStateChange: () => {},
    });
    view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
    const { text } = view.serializeForClipboard(view.state.selection.content());

    expect(text).toBe("bodya picture of a cat7");
    expect(text).not.toContain("Jane Doe");
    expect(text).not.toContain("what the comment says");
    expect(text).not.toContain("what the footnote says");
    view.destroy();
  });

  it("pasting copied HTML keeps the paragraph style", () => {
    const source = openStyledEditor();
    source.view.dispatch(
      source.view.state.tr.setNodeMarkup(0, null, {
        ...source.view.state.doc.child(0).attrs,
        pPr: '<w:pPr><w:pStyle w:val="DocumentHeading1"/></w:pPr>',
      })
    );
    source.view.dispatch(
      source.view.state.tr.setSelection(new AllSelection(source.view.state.doc))
    );
    const html = copiedHtml(source.view);
    source.view.destroy();
    expect(html).toContain('data-style="DocumentHeading1"');

    const { view } = openStyledEditor();
    paste(view, { "text/plain": "source", "text/html": html });

    expect(styleIdOf(view.state.doc.firstChild?.attrs.pPr)).toBe(
      "DocumentHeading1"
    );
    view.destroy();
  });

  it("copying inside a table carries none of what the table is written with", () => {
    const cell = (text: string, inner?: PMNode) =>
      docxSchema.nodes.tableCell.create(
        {
          tcPr: '<w:tcPr><w:shd w:val="CELL_PROPERTIES"/></w:tcPr>',
          tcAttrs: 'w:id="CELL_ATTRIBUTES"',
        },
        inner ??
          docxSchema.nodes.paragraph.create(null, [docxSchema.text(text)])
      );
    const inner = docxSchema.nodes.table.create(
      { tblPr: '<w:tblPr><w:tblStyle w:val="INNER_TABLE"/></w:tblPr>' },
      [docxSchema.nodes.tableRow.create(null, [cell("deep"), cell("also")])]
    );
    const outer = docxSchema.nodes.table.create(
      {
        tblPr: '<w:tblPr><w:tblStyle w:val="OUTER_TABLE"/></w:tblPr>',
        tblAttrs: 'w:x="TABLE_ATTRIBUTES"',
      },
      [
        docxSchema.nodes.tableRow.create(
          { trPr: '<w:trPr><w:x w:val="ROW_PROPERTIES"/></w:trPr>' },
          [cell("", inner), cell("beside")]
        ),
      ]
    );
    const view = createEditorView({
      mount: document.createElement("div"),
      state: createEditorState(docxSchema.nodes.doc.create(null, [outer])),
      onStateChange: () => {},
    });
    let firstText = -1;
    view.state.doc.descendants((node, position) => {
      if (firstText < 0 && node.isText) firstText = position;
      return true;
    });
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, firstText + 1, firstText + 3)
      )
    );

    // prosemirror-view writes the wrappers a slice is open through into `data-pm-slice` after the
    // serializer has run, which is the one place a whitelist over the drawing cannot reach
    const html = copiedHtml(view);
    for (const written of [
      "CELL_PROPERTIES",
      "CELL_ATTRIBUTES",
      "INNER_TABLE",
      "OUTER_TABLE",
      "TABLE_ATTRIBUTES",
      "ROW_PROPERTIES",
    ]) {
      expect(html).not.toContain(written);
    }
    expect(html).not.toContain("<w:");
    view.destroy();
  });

  it("a link leaves as one and comes back as one", () => {
    const paragraph = docxSchema.nodes.paragraph.create(null, [
      docxSchema.text("the docs", [
        docxSchema.marks.link.create({
          href: "https://example.com/docs",
          linkPrefix: '<w:hyperlink r:id="RELATIONSHIP">',
          key: 3,
        }),
      ]),
    ]);
    const source = createEditorView({
      mount: document.createElement("div"),
      state: createEditorState(docxSchema.nodes.doc.create(null, [paragraph])),
      onStateChange: () => {},
    });
    source.dispatch(
      source.state.tr.setSelection(new AllSelection(source.state.doc))
    );
    const html = copiedHtml(source);
    source.destroy();

    // An anchor is what a link is anywhere else, and the relationship it hung off here is not
    expect(html).toContain(
      '<a class="docx-editor-link" href="https://example.com/docs">'
    );
    expect(html).not.toContain("RELATIONSHIP");

    const { view } = openEditor(makeNumberedDocx(PARAGRAPH));
    selectSource(view);
    paste(view, { "text/plain": "the docs", "text/html": html });
    const links: unknown[] = [];
    view.state.doc.descendants((node) => {
      for (const mark of node.marks) {
        if (mark.type === docxSchema.marks.link) links.push(mark.attrs.href);
      }
      return true;
    });

    expect(links).toContain("https://example.com/docs");
    view.destroy();
  });
});
