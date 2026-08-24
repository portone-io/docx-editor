// @vitest-environment jsdom
import {
  DOMSerializer,
  type Mark,
  type Node as PMNode,
} from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  documentXmlOf,
  makeDocx,
  makeStyledDocx,
} from "../../__testing__/docx";
import { posOfText, runCommand, select } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import {
  type ParagraphAlign,
  type RunFormat,
  toParagraphFormat,
  toRunFormat,
} from "../../model/format";
import { docxSchema } from "../../schema";
import { createEditorState } from "../createEditor";
import { docxKeymap } from "../plugins/keymap";
import {
  activeParagraphAlign,
  activeParagraphStyle,
  setParagraphAlign,
  setParagraphStyle,
} from "./paragraphCommands";

function paragraph(text: string, pPr = ""): string {
  return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

/** The state built out of everything the document told us, the same as `DocxEditor` builds it */
function editorStateOf(doc: PMNode, session: SessionStore): EditorState {
  return createEditorState(doc, {
    styles: session.styles,
    defaults: session.defaults,
    paragraphStyles: session.paragraphStyles,
  });
}

function opened(body: string): { state: EditorState; session: SessionStore } {
  const { doc, session } = importDocx(makeDocx(body));
  return { state: editorStateOf(doc, session), session };
}

/** A state with the caret placed on the first character of the given text */
function at(state: EditorState, text: string): EditorState {
  return select(state, posOfText(state.doc, text));
}

function alignOf(doc: PMNode, index: number): ParagraphAlign | undefined {
  return toParagraphFormat(doc.child(index).attrs.format)?.align;
}

describe("applying paragraph alignment", () => {
  it("writes the alignment onto the selected paragraph", () => {
    const { state, session } = opened(paragraph("Body"));
    const centered = runCommand(at(state, "Body"), setParagraphAlign("center"));

    expect(alignOf(centered.doc, 0)).toBe("center");
    expect(documentXmlOf(centered.doc, session)).toContain(
      '<w:pPr><w:jc w:val="center"/></w:pPr>'
    );
  });

  it("justify goes out to the document as both", () => {
    const { state, session } = opened(paragraph("Body"));
    const justified = runCommand(
      at(state, "Body"),
      setParagraphAlign("justify")
    );

    expect(alignOf(justified.doc, 0)).toBe("justify");
    expect(documentXmlOf(justified.doc, session)).toContain(
      '<w:jc w:val="both"/>'
    );
  });

  it("all the selected paragraphs change together", () => {
    const { state } = opened(paragraph("First") + paragraph("Second"));
    const both = select(
      state,
      posOfText(state.doc, "First"),
      posOfText(state.doc, "Second")
    );
    const right = runCommand(both, setParagraphAlign("right"));

    expect(alignOf(right.doc, 0)).toBe("right");
    expect(alignOf(right.doc, 1)).toBe("right");
  });

  it("does nothing in a paragraph that already has that alignment", () => {
    const { state } = opened(
      paragraph("Body", '<w:pPr><w:jc w:val="both"/></w:pPr>')
    );
    const spot = at(state, "Body");
    expect(setParagraphAlign("justify")(spot)).toBe(false);
    expect(setParagraphAlign("center")(spot)).toBe(true);
  });

  it("reads a paragraph where nobody wrote an alignment as left aligned", () => {
    const { state } = opened(paragraph("Body"));
    const spot = at(state, "Body");
    expect(setParagraphAlign("left")(spot)).toBe(false);
  });

  it("changes only the alignment and leaves the rest of the paragraph formatting alone", () => {
    const pPr =
      '<w:pPr><w:spacing w:line="276" w:lineRule="auto"/>' +
      '<w:jc w:val="both"/></w:pPr>';
    const { state, session } = opened(paragraph("Body", pPr));
    const centered = runCommand(at(state, "Body"), setParagraphAlign("center"));

    expect(documentXmlOf(centered.doc, session)).toContain(
      '<w:pPr><w:spacing w:line="276" w:lineRule="auto"/>' +
        '<w:jc w:val="center"/></w:pPr>'
    );
  });
});

describe("deciding the active alignment", () => {
  it("reports that value when all the selected paragraphs share the same alignment", () => {
    const { state } = opened(
      paragraph("Body", '<w:pPr><w:jc w:val="both"/></w:pPr>')
    );
    expect(activeParagraphAlign(at(state, "Body"))).toEqual({
      kind: "shared",
      align: "justify",
    });
  });

  it("reports left when nobody wrote one down", () => {
    const { state } = opened(paragraph("Body"));
    expect(activeParagraphAlign(at(state, "Body"))).toEqual({
      kind: "shared",
      align: "left",
    });
  });

  it("presses nothing when the alignments are mixed", () => {
    const { state } = opened(
      paragraph("First", '<w:pPr><w:jc w:val="center"/></w:pPr>') +
        paragraph("Second")
    );
    const both = select(
      state,
      posOfText(state.doc, "First"),
      posOfText(state.doc, "Second")
    );
    expect(activeParagraphAlign(both)).toEqual({ kind: "mixed" });
  });

  it("the alignment a style gives counts toward the active decision too", () => {
    const { doc, session } = importDocx(
      makeStyledDocx(
        paragraph("Entry", '<w:pPr><w:pStyle w:val="Item"/></w:pPr>'),
        '<w:style w:styleId="Item">' +
          '<w:pPr><w:jc w:val="center"/></w:pPr></w:style>'
      )
    );
    const state = createEditorState(doc, { styles: session.styles });
    const spot = at(state, "Entry");

    expect(activeParagraphAlign(spot)).toEqual({
      kind: "shared",
      align: "center",
    });
    // Choosing the same value the style gives leaves nothing to do
    expect(setParagraphAlign("center")(spot)).toBe(false);

    // Choosing a different value makes the paragraph write it down itself and beat the style
    const right = runCommand(spot, setParagraphAlign("right"));
    expect(activeParagraphAlign(right)).toEqual({
      kind: "shared",
      align: "right",
    });
    expect(documentXmlOf(right.doc, session)).toContain(
      '<w:pStyle w:val="Item"/><w:jc w:val="right"/>'
    );
  });
});

/** A paragraph style that lays down both paragraph values and run values */
const QUOTE_STYLE =
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>' +
  '<w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>' +
  '<w:rPr><w:i/><w:color w:val="2E74B5"/><w:sz w:val="32"/></w:rPr></w:style>';

/** A run that writes down a color of its own, so the layering can be told apart */
const RED_RUN =
  '<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr>' +
  '<w:t xml:space="preserve">Quoted</w:t></w:r>';

function openedStyled(
  body: string,
  styles = QUOTE_STYLE
): { state: EditorState; session: SessionStore } {
  const { doc, session } = importDocx(makeStyledDocx(body, styles));
  return { state: editorStateOf(doc, session), session };
}

/** The run mark covering the given text. null when the text carries none */
function markOf(doc: PMNode, needle: string): Mark | null {
  let found: Mark | null = null;
  doc.descendants((node) => {
    if (!node.isText || node.text !== needle) return true;
    found = node.marks.find((entry) => entry.type.name === "run") ?? null;
    return false;
  });
  return found;
}

/** The display values of the run mark covering the given text */
function markFormatOf(doc: PMNode, needle: string): RunFormat | null {
  return toRunFormat(markOf(doc, needle)?.attrs.format);
}

describe("applying a paragraph style", () => {
  it("writes the style name and lays its values under the paragraph and the text", () => {
    const { state, session } = openedStyled(`<w:p>${RED_RUN}</w:p>`);
    const quoted = runCommand(at(state, "Quoted"), setParagraphStyle("Quote"));

    expect(documentXmlOf(quoted.doc, session)).toContain(
      '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>'
    );
    expect(toParagraphFormat(quoted.doc.child(0).attrs.format)).toEqual({
      align: "center",
      spaceAfterPt: 12,
    });
    // The style is laid underneath, and the color the run wrote down stays on top
    expect(markFormatOf(quoted.doc, "Quoted")).toEqual({
      italic: true,
      color: "#FF0000",
      fontSizePt: 16,
    });
  });

  it("the style name leads the formatting the paragraph already wrote down", () => {
    const { state, session } = openedStyled(
      paragraph("Body", '<w:pPr><w:jc w:val="both"/></w:pPr>')
    );
    const quoted = runCommand(at(state, "Body"), setParagraphStyle("Quote"));

    expect(documentXmlOf(quoted.doc, session)).toContain(
      '<w:pPr><w:pStyle w:val="Quote"/><w:jc w:val="both"/></w:pPr>'
    );
    // What the paragraph wrote down still beats the style
    expect(toParagraphFormat(quoted.doc.child(0).attrs.format)).toEqual({
      align: "justify",
      spaceAfterPt: 12,
    });
  });

  it("clearing the style takes the name away and the values with it", () => {
    const { state, session } = openedStyled(
      `<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>${RED_RUN}</w:p>`
    );
    const plain = runCommand(at(state, "Quoted"), setParagraphStyle(null));

    expect(documentXmlOf(plain.doc, session)).not.toContain("<w:pStyle");
    expect(toParagraphFormat(plain.doc.child(0).attrs.format)).toBeNull();
    // Only what the run itself wrote down is left
    expect(markFormatOf(plain.doc, "Quoted")).toEqual({ color: "#FF0000" });
  });

  it("all the selected paragraphs change together", () => {
    const { state } = openedStyled(paragraph("First") + paragraph("Second"));
    const both = select(
      state,
      posOfText(state.doc, "First"),
      posOfText(state.doc, "Second")
    );
    const quoted = runCommand(both, setParagraphStyle("Quote"));

    expect(quoted.doc.child(0).attrs.pPr).toContain(
      '<w:pStyle w:val="Quote"/>'
    );
    expect(quoted.doc.child(1).attrs.pPr).toContain(
      '<w:pStyle w:val="Quote"/>'
    );
  });

  it("does nothing in a paragraph that already points at that style", () => {
    const { state } = openedStyled(
      paragraph("Body", '<w:pPr><w:pStyle w:val="Quote"/></w:pPr>')
    );
    const spot = at(state, "Body");
    expect(setParagraphStyle("Quote")(spot)).toBe(false);
    expect(setParagraphStyle(null)(spot)).toBe(true);
  });

  it("does nothing in a paragraph that already wears the default style", () => {
    const { state } = openedStyled(paragraph("Body"));
    expect(setParagraphStyle(null)(at(state, "Body"))).toBe(false);
  });

  // Marking the locked text would have the guard refuse the plain paragraph along with it
  it("points a paragraph holding a lock at the style and leaves its text alone", () => {
    const locked =
      '<w:sdt><w:sdtPr><w:id w:val="7"/>' +
      '<w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>' +
      '<w:r><w:t xml:space="preserve">Locked</w:t></w:r>' +
      "</w:sdtContent></w:sdt>";
    const { state } = openedStyled(`<w:p>${locked}</w:p>${paragraph("Body")}`);
    const both = select(
      state,
      posOfText(state.doc, "Locked"),
      posOfText(state.doc, "Body")
    );
    const quoted = runCommand(both, setParagraphStyle("Quote"));

    expect(quoted.doc.child(0).attrs.pPr).toContain(
      '<w:pStyle w:val="Quote"/>'
    );
    expect(quoted.doc.child(1).attrs.pPr).toContain(
      '<w:pStyle w:val="Quote"/>'
    );
    expect(markFormatOf(quoted.doc, "Locked")).toBeNull();
    expect(markFormatOf(quoted.doc, "Body")).toEqual({
      italic: true,
      color: "#2E74B5",
      fontSizePt: 16,
    });
  });
});

/** A paragraph style laying down paragraph values only, with nothing for the text */
const LEAD_STYLE =
  '<w:style w:type="paragraph" w:styleId="Lead"><w:name w:val="Lead"/>' +
  '<w:pPr><w:jc w:val="right"/></w:pPr></w:style>';

/** A paragraph the way the editor builds it while typing: text carrying no run mark */
function typedDoc(text: string): PMNode {
  return docxSchema.nodes.doc.create(null, [
    docxSchema.nodes.paragraph.create(
      { srcId: null, pAttrs: null, pPr: null, format: null },
      docxSchema.text(text)
    ),
  ]);
}

/** An editor holding freshly typed text, with the styles of a document that defines them */
function typed(text: string): { state: EditorState; session: SessionStore } {
  const { session } = importDocx(
    makeStyledDocx(paragraph("Body"), QUOTE_STYLE + LEAD_STYLE)
  );
  return { state: editorStateOf(typedDoc(text), session), session };
}

describe("applying a paragraph style to text typed in the editor", () => {
  it("marks text carrying no run mark with the values the style gives", () => {
    const { state, session } = typed("draft");
    const quoted = runCommand(at(state, "draft"), setParagraphStyle("Quote"));

    expect(markFormatOf(quoted.doc, "draft")).toEqual({
      italic: true,
      color: "#2E74B5",
      fontSizePt: 16,
    });
    // The mark writes down no run formatting of its own, so the text goes out in a bare run
    expect(markOf(quoted.doc, "draft")?.attrs).toMatchObject({
      rPr: null,
      rAttrs: null,
    });
    expect(documentXmlOf(quoted.doc, session)).toContain(
      '<w:r><w:t xml:space="preserve">draft</w:t></w:r>'
    );
  });

  it("leaves the text unmarked when the style lays down no run formatting", () => {
    const { state } = typed("draft");
    const led = runCommand(at(state, "draft"), setParagraphStyle("Lead"));

    expect(markOf(led.doc, "draft")).toBeNull();
    expect(toParagraphFormat(led.doc.child(0).attrs.format)).toEqual({
      align: "right",
    });
  });
});

describe("deciding the active paragraph style", () => {
  it("reports the name when all the selected paragraphs point at the same one", () => {
    const { state } = openedStyled(
      paragraph("Body", '<w:pPr><w:pStyle w:val="Quote"/></w:pPr>')
    );
    expect(activeParagraphStyle(at(state, "Body"))).toEqual({
      kind: "shared",
      styleId: "Quote",
    });
  });

  it("reports no style for a paragraph wearing the default one", () => {
    const { state } = openedStyled(paragraph("Body"));
    expect(activeParagraphStyle(at(state, "Body"))).toEqual({
      kind: "shared",
      styleId: null,
    });
  });

  it("reports them as mixed when the paragraphs point at different ones", () => {
    const { state } = openedStyled(
      paragraph("First", '<w:pPr><w:pStyle w:val="Quote"/></w:pPr>') +
        paragraph("Second")
    );
    const both = select(
      state,
      posOfText(state.doc, "First"),
      posOfText(state.doc, "Second")
    );
    expect(activeParagraphStyle(both)).toEqual({ kind: "mixed" });
  });

  // A selection holding no paragraph has no style to show, and the picker is turned off over it
  it("reports none where the selection holds no paragraph", () => {
    const { doc, session } = importDocx(makeDocx("<w:customXml/>"));
    const state = editorStateOf(doc, session);

    expect(state.doc.child(0).type.name).toBe("docxRaw");
    expect(activeParagraphStyle(state)).toEqual({ kind: "none" });
  });
});

/** The style every paragraph pointing at none of its own wears, which OOXML marks `w:default="1"` */
const NORMAL_STYLE =
  '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
  '<w:name w:val="Normal"/>' +
  '<w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>' +
  '<w:rPr><w:sz w:val="22"/></w:rPr></w:style>';

const HEADING_STYLE =
  '<w:style w:type="paragraph" w:styleId="Heading1">' +
  '<w:name w:val="heading 1"/><w:pPr><w:jc w:val="left"/></w:pPr>' +
  '<w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:style>';

/** The values the paragraph and the character formatting of its style are drawn with */
function displayValuesOf(doc: PMNode, index: number) {
  const attrs = doc.child(index).attrs;
  return {
    format: toParagraphFormat(attrs.format),
    styleRun: toRunFormat(attrs.styleRun),
  };
}

describe("the default paragraph style", () => {
  it("is worn by a paragraph that points at no style of its own", () => {
    const { state } = openedStyled(paragraph("Body"), NORMAL_STYLE);

    expect(displayValuesOf(state.doc, 0)).toEqual({
      format: { align: "center", spaceAfterPt: 12 },
      styleRun: { fontSizePt: 11 },
    });
    expect(markFormatOf(state.doc, "Body")).toEqual({ fontSizePt: 11 });
  });

  it("comes back when the style a paragraph named is cleared", () => {
    const { state, session } = openedStyled(
      paragraph("Heading", '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'),
      NORMAL_STYLE + HEADING_STYLE
    );
    expect(displayValuesOf(state.doc, 0)).toEqual({
      format: { align: "left" },
      styleRun: { bold: true, fontSizePt: 20 },
    });

    const plain = runCommand(at(state, "Heading"), setParagraphStyle(null));
    expect(displayValuesOf(plain.doc, 0)).toEqual({
      format: { align: "center", spaceAfterPt: 12 },
      styleRun: { fontSizePt: 11 },
    });
    expect(documentXmlOf(plain.doc, session)).not.toContain("<w:pStyle");
  });

  // Naming the default style outright is the same as wearing it by inheritance
  it("is applied when it is named outright as well", () => {
    const { state, session } = openedStyled(
      paragraph("Heading", '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'),
      NORMAL_STYLE + HEADING_STYLE
    );
    const normal = runCommand(
      at(state, "Heading"),
      setParagraphStyle("Normal")
    );

    expect(displayValuesOf(normal.doc, 0)).toEqual({
      format: { align: "center", spaceAfterPt: 12 },
      styleRun: { fontSizePt: 11 },
    });
    expect(documentXmlOf(normal.doc, session)).toContain(
      '<w:pStyle w:val="Normal"/>'
    );
  });

  it("stays laid underneath after the paragraph is edited", () => {
    const { state } = openedStyled(paragraph("Body"), NORMAL_STYLE);
    const right = runCommand(at(state, "Body"), setParagraphAlign("right"));

    // The alignment the paragraph now writes down beats the style, and the rest of it stands
    expect(displayValuesOf(right.doc, 0)).toEqual({
      format: { align: "right", spaceAfterPt: 12 },
      styleRun: { fontSizePt: 11 },
    });
  });
});

/** The paragraph as it is drawn on screen */
function renderedParagraph(doc: PMNode, index: number): string {
  const host = document.createElement("div");
  host.appendChild(
    DOMSerializer.fromSchema(docxSchema).serializeNode(doc.child(index))
  );
  return host.innerHTML;
}

/**
 * Text typed into a paragraph carries no run mark of its own, so it is the paragraph that has to
 * draw it in the character formatting of the style it wears. Otherwise it is drawn plain and turns
 * into styled text only after the document is saved and opened again.
 */
describe("typing in a paragraph that wears a style", () => {
  it("draws text typed into an empty paragraph in the style's character formatting", () => {
    const { state, session } = openedStyled("<w:p/>");
    const quoted = runCommand(select(state, 1), setParagraphStyle("Quote"));
    const typedIn = quoted.apply(quoted.tr.insertText("draft", 1));

    expect(markOf(typedIn.doc, "draft")).toBeNull();
    expect(toRunFormat(typedIn.doc.child(0).attrs.styleRun)).toEqual({
      italic: true,
      color: "#2E74B5",
      fontSizePt: 16,
    });
    const drawn = renderedParagraph(typedIn.doc, 0);
    expect(drawn).toContain("font-style: italic");
    expect(drawn).toContain("color: rgb(46, 116, 181)");
    expect(drawn).toContain("font-size: 16pt");

    // None of it reaches the document: the paragraph names the style and the text goes out bare
    const documentXml = documentXmlOf(typedIn.doc, session);
    expect(documentXml).toContain(
      '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>' +
        '<w:r><w:t xml:space="preserve">draft</w:t></w:r></w:p>'
    );
    expect(documentXml).not.toContain("styleRun");
  });

  it("draws text typed into the paragraph Enter opened the same way", () => {
    const { state } = openedStyled(
      paragraph("Quoted", '<w:pPr><w:pStyle w:val="Quote"/></w:pPr>')
    );
    // The caret at the very end of the paragraph, which is where Enter opens an empty one
    const withCaret = select(state, state.doc.child(0).nodeSize - 1);
    let split = withCaret;
    expect(
      docxKeymap.Enter(withCaret, (tr) => {
        split = withCaret.apply(tr);
      })
    ).toBe(true);
    const typedIn = split.apply(
      split.tr.insertText("draft", split.selection.head)
    );

    expect(markOf(typedIn.doc, "draft")).toBeNull();
    expect(toRunFormat(typedIn.doc.child(1).attrs.styleRun)).toEqual({
      italic: true,
      color: "#2E74B5",
      fontSizePt: 16,
    });
    expect(renderedParagraph(typedIn.doc, 1)).toContain("font-style: italic");
  });
});
