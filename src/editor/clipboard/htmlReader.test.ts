// @vitest-environment jsdom

import { AllSelection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeNumberedDocx,
  makeStyledDocx,
  makeStyledNumberedDocx,
} from "../../__testing__/docx";
import { exportDocx } from "../../docx/exportDocx";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { toParagraphFormat, toRunFormat } from "../../model/format";
import { newListsOf } from "../../numbering/listRegistry";
import { templateList } from "../../numbering/listTemplate";
import { styleIdOf } from "../../ooxml/props";
import { editorClassNames } from "../../styles/classNames";
import { listRefOf } from "../commands/listCommands";
import { createEditorView, editorStateForSession } from "../createEditor";
import {
  documentNumbering,
  paragraphMarkers,
} from "../plugins/numberingDecorations";

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
    state: editorStateForSession({ doc, session }),
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
    state: editorStateForSession({ doc, session }),
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

describe("pasting supported HTML", () => {
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

  it("registers a definition for each list it pastes, as starting one does", () => {
    const { view } = openEditor();
    view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
    paste(view, {
      "text/plain": "One\nTwo\nNested",
      "text/html": "<ol><li>One</li><li>Two<ul><li>Nested</li></ul></li></ol>",
    });

    const registered = newListsOf(view.state.doc.attrs.newLists);
    expect([...registered.keys()].sort()).toEqual([2, 3]);
    expect(registered.get(2)).toEqual(templateList("numbered"));
    expect(registered.get(3)).toEqual(templateList("bullet"));
    // The numbers drawn come from those definitions and not from the ids they were given: the
    // nested item stands at the second level of the bullet list it started
    expect(
      paragraphMarkers(view.state.doc, documentNumbering(view.state)).map(
        (marker) => marker.text
      )
    ).toEqual(["1.", "2.", "○"]);
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

  it("a pasted list paragraph carries the default style's character values", () => {
    const styles =
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
      '<w:name w:val="Normal"/><w:pPr><w:spacing w:after="160"/></w:pPr>' +
      '<w:rPr><w:b/><w:sz w:val="22"/></w:rPr></w:style>';
    const { doc, session } = importDocx(
      makeStyledNumberedDocx(
        '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>',
        styles
      )
    );
    const view = createEditorView({
      mount: document.createElement("div"),
      state: editorStateForSession({ doc, session }),
      onStateChange: () => {},
    });
    view.dispatch(view.state.tr.setSelection(new AllSelection(view.state.doc)));
    paste(view, {
      "text/plain": "One\nTwo",
      "text/html": "<ol><li>One</li><li>Two</li></ol>",
    });

    const item = view.state.doc.child(0);
    expect(listRefOf(item)).toEqual({ numId: 2, ilvl: 0 });
    expect(toParagraphFormat(item.attrs.format)).toEqual({
      numbering: { numId: 2, ilvl: 0 },
      spaceAfterPt: 8,
      // The list it joined hangs its number, and that is where the text of the item begins
      tabStops: [{ positionPt: 36, align: "start" }],
    });
    expect(toRunFormat(item.attrs.styleRun)).toEqual({
      bold: true,
      fontSizePt: 11,
    });
    expect(documentXmlOf(view.state.doc, session)).toContain(
      '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr>'
    );
    view.destroy();
  });
});
