// @vitest-environment jsdom
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

  it("keeps the size of an image copied inside the editor", () => {
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

    // The measure the document keeps does not travel; the pixels the browser was given do
    expect(html).not.toContain("data-extent");
    expect(html).toContain(`width="${Math.round(emuToPx(extent.cx))}"`);

    paste(view, { "text/plain": "", "text/html": html });
    const pasted = view.state.doc.child(0).child(0);
    expect(pasted.type.name).toBe("image");
    expect(pasted.attrs.extent).toEqual(extent);
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
