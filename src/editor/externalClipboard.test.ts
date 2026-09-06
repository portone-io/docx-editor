// @vitest-environment jsdom

import type { Node as PMNode } from "prosemirror-model";
import { AllSelection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeNumberedDocx,
  makeStyledDocx,
  TINY_PNG_DATA_URL,
} from "../__testing__/docx";
import { exportDocx } from "../docx/exportDocx";
import { NO_DOCUMENT_DEFAULTS, styleIdOf } from "../docx/formatting";
import { importDocx } from "../docx/importDocx";
import type { SessionStore } from "../docx/session";
import { toRunFormat } from "../model/format";
import { parseNumbering } from "../numbering/parseNumbering";
import { emuToPx } from "../ooxml/image";
import { docxSchema } from "../schema";
import { editorClassNames } from "../styles/classNames";
import { listRefOf } from "./commands/listCommands";
import { createEditorState, createEditorView } from "./createEditor";

function openEditor(canStartNewList = true): {
  view: EditorView;
  session: SessionStore;
} {
  const body = '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>';
  const { doc, session } = importDocx(
    canStartNewList ? makeNumberedDocx(body) : makeDocx(body)
  );
  const view = createEditorView({
    mount: document.createElement("div"),
    state: createEditorState(doc, {
      numbering: parseNumbering(session.numberingXml),
      canStartNewList,
    }),
    defaults: NO_DOCUMENT_DEFAULTS,
    onStateChange: () => {},
  });
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
  );
  return { view, session };
}

function openStyledEditor(): { view: EditorView; session: SessionStore } {
  const body = '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>';
  const styles =
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
    '<w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="DocumentHeading1">' +
    '<w:name w:val="heading 1"/><w:pPr><w:spacing w:before="240"/></w:pPr>' +
    '<w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>';
  const { doc, session } = importDocx(makeStyledDocx(body, styles));
  const view = createEditorView({
    mount: document.createElement("div"),
    state: createEditorState(doc, {
      styles: session.styles,
      defaults: session.defaults,
      paragraphStyles: session.paragraphStyles,
      canStartNewList: false,
    }),
    defaults: session.defaults,
    onStateChange: () => {},
  });
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
  );
  return { view, session };
}

function paste(view: EditorView, data: Record<string, string>): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => data[type] ?? "" },
  });
  view.dom.dispatchEvent(event);
}

function runFormat(view: EditorView, text: string) {
  let found = null;
  view.state.doc.descendants((node) => {
    if (!node.isText || node.text !== text) return true;
    const mark = node.marks.find((candidate) => candidate.type.name === "run");
    found = toRunFormat(mark?.attrs.format);
    return false;
  });
  return found;
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
    defaults: NO_DOCUMENT_DEFAULTS,
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
    defaults: NO_DOCUMENT_DEFAULTS,
    onStateChange: () => {},
  });
  view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
  const { text } = view.serializeForClipboard(view.state.selection.content());
  view.destroy();
  return text;
}

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
      defaults: NO_DOCUMENT_DEFAULTS,
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
      defaults: NO_DOCUMENT_DEFAULTS,
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
      defaults: NO_DOCUMENT_DEFAULTS,
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

  it("a copy taken before the style travelled on its own still keeps it", () => {
    const { view } = openStyledEditor();
    paste(view, {
      "text/plain": "source",
      "text/html":
        `<p class="${editorClassNames.paragraph}" ` +
        "data-ppr='<w:pPr><w:pStyle w:val=\"DocumentHeading1\"/></w:pPr>'>" +
        "source</p>",
    });

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
      defaults: NO_DOCUMENT_DEFAULTS,
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
          linkKey: 3,
        }),
      ]),
    ]);
    const source = createEditorView({
      mount: document.createElement("div"),
      state: createEditorState(docxSchema.nodes.doc.create(null, [paragraph])),
      defaults: NO_DOCUMENT_DEFAULTS,
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

    const { view } = openEditor();
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

describe("pasting supported HTML", () => {
  it("maps an HTML heading to the destination document heading style", () => {
    const { view, session } = openStyledEditor();
    paste(view, {
      "text/plain": "Heading",
      "text/html": "<h1>Heading</h1>",
    });

    const heading = view.state.doc.firstChild;
    expect(styleIdOf(heading?.attrs.pPr)).toBe("DocumentHeading1");
    expect(heading?.attrs.styleRun).toMatchObject({
      bold: true,
      fontSizePt: 20,
    });
    expect(runFormat(view, "Heading")).toBeNull();
    const reopened = importDocx(exportDocx(view.state.doc, session)).doc;
    expect(styleIdOf(reopened.firstChild?.attrs.pPr)).toBe("DocumentHeading1");
    view.destroy();
  });

  it("keeps a copied editor heading semantic instead of baking its display size", () => {
    const { view } = openStyledEditor();
    paste(view, {
      "text/plain": "Copied heading",
      "text/html":
        `<p class="${editorClassNames.paragraph}" ` +
        'data-pm-slice="0 0 []" ' +
        'data-style="DocumentHeading1" ' +
        'style="font-size: 26pt; font-weight: 700">Copied heading</p>',
    });

    expect(styleIdOf(view.state.doc.firstChild?.attrs.pPr)).toBe(
      "DocumentHeading1"
    );
    expect(runFormat(view, "Copied heading")).toBeNull();
    view.destroy();
  });

  it("keeps an open editor slice inline without changing the destination paragraph style", () => {
    const { view } = openStyledEditor();
    paste(view, {
      "text/plain": "Partial heading",
      "text/html":
        `<p class="${editorClassNames.paragraph}" ` +
        'data-pm-slice="1 1 []" ' +
        'data-style="DocumentHeading1" ' +
        'style="font-size: 26pt; font-weight: 700">Partial heading</p>',
    });

    expect(styleIdOf(view.state.doc.firstChild?.attrs.pPr)).toBeNull();
    expect(runFormat(view, "Partial heading")).toMatchObject({
      bold: true,
      fontSizePt: 26,
    });
    view.destroy();
  });

  it("keeps copied editor tabs without trusting private HTML attributes", () => {
    const { view, session } = openEditor();
    paste(view, {
      "text/plain": "A\tB",
      "text/html":
        '<p data-pm-slice="0 0 []">A<span class="' +
        editorClassNames.tab +
        '" data-tattrs="x=&quot;/>&lt;w:t>injected&lt;/w:t>&lt;w:tab x=&quot;">\t</span>B</p>',
    });

    expect(view.state.doc.textContent).toBe("A\tB");
    const xml = documentXmlOf(view.state.doc, session);
    expect(xml).toContain("<w:tab/>");
    expect(xml).not.toContain("injected");
    view.destroy();
  });

  it("ignores malformed copied image metadata without failing the paste", () => {
    const { view } = openEditor();
    paste(view, {
      "text/plain": "Before After",
      "text/html":
        '<p data-pm-slice="0 0 []">Before ' +
        `<img class="${editorClassNames.image}" src="javascript:alert(1)" ` +
        'data-extent="not json"> After</p>',
    });

    expect(view.state.doc.firstChild?.textContent).toBe("Before  After");
    expect(view.state.doc.firstChild?.childCount).toBe(1);
    view.destroy();
  });

  it("uses direct heading formatting when the destination defines no heading style", () => {
    const { view } = openEditor(false);
    paste(view, {
      "text/plain": "Fallback heading",
      "text/html": "<h2>Fallback heading</h2>",
    });

    expect(styleIdOf(view.state.doc.firstChild?.attrs.pPr)).toBeNull();
    expect(runFormat(view, "Fallback heading")).toMatchObject({
      bold: true,
      fontSizePt: 18,
    });
    view.destroy();
  });

  it("writes supported font formatting into DOCX run properties", () => {
    const { view, session } = openEditor();
    paste(view, {
      "text/plain": "Styled",
      "text/html":
        '<p><span style="font-family: Arial; font-size: 16pt; font-weight: 700; font-style: italic; text-decoration: underline line-through; color: #123456; background-color: #abcdef">Styled</span></p>',
    });

    expect(runFormat(view, "Styled")).toMatchObject({
      bold: true,
      italic: true,
      underline: "single",
      strike: true,
      fontSizePt: 16,
      fontFamily: '"Arial"',
      color: "#123456",
      background: "#ABCDEF",
    });
    const xml = documentXmlOf(view.state.doc, session);
    expect(xml).toContain('<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>');
    expect(xml).toContain('<w:sz w:val="32"/>');
    expect(xml).toContain("<w:b/>");
    expect(xml).toContain("<w:i/>");
    expect(xml).toContain("<w:strike/>");
    expect(xml).toContain('<w:u w:val="single"/>');
    expect(xml).toContain('<w:color w:val="123456"/>');
    expect(xml).toContain('w:fill="ABCDEF"');
    const reopened = importDocx(exportDocx(view.state.doc, session)).doc;
    const reopenedRun = reopened.firstChild?.firstChild?.marks.find(
      (mark) => mark.type.name === "run"
    );
    expect(toRunFormat(reopenedRun?.attrs.format)).toMatchObject({
      bold: true,
      italic: true,
      underline: "single",
      strike: true,
      fontSizePt: 16,
      fontFamily: '"Arial"',
      color: "#123456",
      background: "#ABCDEF",
    });
    view.destroy();
  });

  it("keeps paragraphs, breaks, safe links and basic list structure", () => {
    const { view, session } = openEditor();
    view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
    paste(view, {
      "text/plain": "Intro\nline link\nOne\nTwo\nNested",
      "text/html":
        '<p>Intro<br>line <a href="https://example.com/docs">link</a></p>' +
        "<ol><li>One</li><li>Two<ul><li>Nested</li></ul></li></ol>",
    });

    expect(view.state.doc.childCount).toBe(4);
    expect(view.state.doc.child(0).textContent).toBe("Introline link");
    expect(view.state.doc.child(0).child(1).type.name).toBe("hardBreak");
    expect(listRefOf(view.state.doc.child(1))).toEqual({ numId: 2, ilvl: 0 });
    expect(listRefOf(view.state.doc.child(2))).toEqual({ numId: 2, ilvl: 0 });
    expect(listRefOf(view.state.doc.child(3))).toEqual({ numId: 3, ilvl: 1 });
    const link = view.state.doc
      .child(0)
      .child(3)
      .marks.find((mark) => mark.type.name === "link");
    expect(link?.attrs.href).toBe("https://example.com/docs");
    const xml = documentXmlOf(view.state.doc, session);
    expect(xml).toContain('<w:numId w:val="2"/>');
    expect(xml).toContain('<w:numId w:val="3"/>');
    expect(xml).toContain("<w:hyperlink");
    view.destroy();
  });

  it("drops unsafe links and keeps their readable text", () => {
    const { view } = openEditor();
    paste(view, {
      "text/plain": "Open",
      "text/html": '<p><a href="javascript:alert(1)">Open</a></p>',
    });

    const text = view.state.doc.firstChild?.firstChild;
    expect(text?.textContent).toBe("Open");
    expect(text?.marks.some((mark) => mark.type.name === "link")).toBe(false);
    view.destroy();
  });

  it("keeps list items readable when the document cannot create numbering rules", () => {
    const { view } = openEditor(false);
    paste(view, {
      "text/plain": "One\nTwo",
      "text/html": "<ul><li>One</li><li>Two</li></ul>",
    });

    expect(view.state.doc.textContent).toBe("• One• Two");
    expect(listRefOf(view.state.doc.child(0))).toBeNull();
    expect(listRefOf(view.state.doc.child(1))).toBeNull();
    view.destroy();
  });
});
