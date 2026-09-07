// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { afterEach, describe, expect, it } from "vitest";
import {
  bytesEqual,
  makeNumberedDocx,
  readFixture,
} from "../../__testing__/docx";
import { exportDocx } from "../../docx/exportDocx";
import { NO_FORMATTING, readRunFormat } from "../../docx/formatting";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { toParagraphFormat, toRunFormat } from "../../model/format";
import { type Numbering, parseNumbering } from "../../numbering/parseNumbering";
import { docxSchema } from "../../schema";
import { editorAttributes, editorCssVariables } from "../../styles/classNames";
import {
  createEditorState,
  createEditorView,
  editorStateForSession,
} from "../createEditor";
import { type EditorDocument, NO_DOCUMENT } from "../editorDocument";
import {
  markerDecorations,
  type PlacedMarker,
  paragraphMarkers,
} from "./numberingDecorations";

/** The fixture with the lists, both in the body and inside a table cell */
const LIST_FIXTURE = "kitchen-sink.docx";

/** The fixture whose list paragraphs write their own indent down */
const INDENTED_FIXTURE = "size-fallback.docx";

function openWithNumbering(name: string) {
  const bytes = readFixture(name);
  const { doc, session } = importDocx(bytes);
  return {
    bytes,
    doc,
    session,
    numbering: parseNumbering(session.numberingXml),
  };
}

function markersOf(name: string) {
  const { doc, numbering } = openWithNumbering(name);
  return paragraphMarkers(doc, numbering);
}

const mounted: { view: { destroy: () => void } | null } = { view: null };

afterEach(() => {
  mounted.view?.destroy();
  mounted.view = null;
});

describe("list marker decorations", () => {
  it("every list paragraph gets a marker that is not blank", () => {
    const markers = markersOf(LIST_FIXTURE);
    expect(markers.length).toBeGreaterThan(0);
    expect(markers.every((marker) => marker.text.length > 0)).toBe(true);
  });

  it("the first list's numbers run on from 1", () => {
    expect(
      markersOf(LIST_FIXTURE)
        .slice(0, 5)
        .map((m) => m.text)
    ).toEqual(["1.", "1)", "2)", "3)", "4)"]);
  });

  it("attaches nothing to a document with no lists", () => {
    expect(markersOf("east-asian.docx")).toEqual([]);
  });

  it("takes the width the marker sits in from the paragraph's hanging indent", () => {
    // The first list paragraph has w:ind hanging="400" = 20pt
    expect(markersOf(LIST_FIXTURE)[0].widthPt).toBe(20);
  });

  it("there is exactly one decoration per paragraph", () => {
    const { doc, numbering } = openWithNumbering(LIST_FIXTURE);
    expect(markerDecorations(doc, numbering).find()).toHaveLength(
      paragraphMarkers(doc, numbering).length
    );
  });
});

describe("decorations do not touch the document", () => {
  it.each([LIST_FIXTURE, INDENTED_FIXTURE])(
    "%s: exporting while the markers are showing is byte identical to the original",
    (name) => {
      const { bytes, doc, session } = openWithNumbering(name);
      const out = exportDocx(
        editorStateForSession({ doc, session }).doc,
        session
      );

      const original = unzipSync(bytes);
      const exported = unzipSync(out);
      for (const key of Object.keys(original)) {
        expect(bytesEqual(exported[key], original[key])).toBe(true);
      }
    }
  );

  it("the markers do not mix into the document's text", () => {
    const { doc, session } = openWithNumbering(LIST_FIXTURE);
    const withMarkers = editorStateForSession({ doc, session });
    const withoutMarkers = createEditorState(doc);
    expect(withMarkers.doc.textContent).toBe(withoutMarkers.doc.textContent);
  });

  it("recounts the numbers that follow when a paragraph disappears", () => {
    const { doc, session, numbering } = openWithNumbering(LIST_FIXTURE);
    const state = editorStateForSession({ doc, session });
    const before = paragraphMarkers(state.doc, numbering);

    const removed = state.apply(state.tr.delete(before[0].from, before[0].to));
    const after = paragraphMarkers(removed.doc, numbering);

    expect(after.length).toBe(before.length - 1);
    expect(after[0].text).toBe("1)");
  });
});

/**
 * Where Word places list paragraphs.
 *
 * The number is left-aligned at where the hanging indent begins (left - hanging), and the body
 * text starts at left. The slot the number sits in is the same as the hanging width, so the gap
 * between the number and the body is only that width minus the length of the number.
 */
describe("where list numbers sit (Word geometry)", () => {
  const toTwips = (pt: number) => Math.round(pt * 20);

  /** Where one numbered paragraph is placed on screen. The unit is twips */
  function geometry(doc: PMNode, marker: PlacedMarker) {
    const format = toParagraphFormat(doc.nodeAt(marker.from)?.attrs.format);
    // What the paragraph wrote down wins; otherwise the value the level gives is what applies on screen
    const startPt = format?.indentStartPt ?? marker.indentStartPt ?? 0;
    const textIndentPt = format?.textIndentPt ?? marker.textIndentPt ?? 0;
    return {
      body: toTwips(startPt),
      marker: toTwips(startPt + textIndentPt),
      width: toTwips(marker.widthPt),
    };
  }

  function geometryOfText(name: string, needle: string) {
    const { doc, numbering } = openWithNumbering(name);
    const found = paragraphMarkers(doc, numbering).find((marker) =>
      doc.nodeAt(marker.from)?.textContent.includes(needle)
    );
    if (!found) throw new Error(`numbered paragraph not found: ${needle}`);
    return geometry(doc, found);
  }

  it("uses the ind the paragraph wrote down as is (left=708 hanging=360)", () => {
    expect(
      geometryOfText(INDENTED_FIXTURE, "And she tried to curtsey")
    ).toEqual({ body: 708, marker: 348, width: 360 });
  });

  it("a sub item follows the same rule (left=1133 hanging=360)", () => {
    expect(geometryOfText(INDENTED_FIXTURE, "by dropping a curtsey")).toEqual({
      body: 1133,
      marker: 773,
      width: 360,
    });
  });
});

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/**
 * A list of one level, holding whatever properties the test writes into it, read the way an
 * opened document reads its own numbering part
 */
function oneLevel(properties: string): Numbering {
  return parseNumbering(
    `<w:numbering ${W_NS}><w:abstractNum w:abstractNumId="0">` +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
      `${properties}<w:lvlText w:val="%1."/>` +
      "</w:lvl></w:abstractNum>" +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>',
    { readRun: (rPr) => readRunFormat(rPr) }
  );
}

function listDoc(
  format: Record<string, unknown>,
  pPr: string | null = null
): PMNode {
  return docxSchema.nodes.doc.create(null, [
    docxSchema.nodes.paragraph.create(
      { srcId: "opened:body:0", pAttrs: null, pPr, format },
      docxSchema.text("Item")
    ),
  ]);
}

const numbered = { numbering: { numId: 1, ilvl: 0 } };

/** The properties the display values above are worked out from, which a state reads them off again */
const NUMBERED_PPR =
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';

/** A document that knows nothing but these level definitions */
const knowing = (numbering: Numbering): EditorDocument => ({
  ...NO_DOCUMENT,
  formatting: { ...NO_FORMATTING, numbering },
});

/** The paragraph as the editor draws it, with the marker decoration applied */
function drawnParagraph(numbering: Numbering): HTMLElement {
  const mount = document.createElement("div");
  document.body.appendChild(mount);
  mounted.view = createEditorView({
    mount,
    state: createEditorState(listDoc(numbered, NUMBERED_PPR), {
      document: knowing(numbering),
    }),
    onStateChange: () => {},
  });
  const paragraph = mount.querySelector("p.docx-editor-p[data-marker]");
  if (!(paragraph instanceof HTMLElement)) throw new Error("no paragraph");
  return paragraph;
}

/**
 * When a paragraph has not written down a `w:ind`, Word uses the indentation the list level
 * specifies. Without overlaying that value on screen, the number would sit at the paragraph's
 * left edge and the whole body text would be pushed over by the number's width.
 */
describe("the indentation the level specifies", () => {
  const numbering: Numbering = oneLevel(
    '<w:pPr><w:ind w:left="720" w:right="240" w:hanging="360"/></w:pPr>'
  );

  const withLevels: EditorDocument = knowing(numbering);

  it("overlays the level's indentation on screen when the paragraph has no ind", () => {
    const [marker] = paragraphMarkers(listDoc(numbered), numbering);
    expect(marker).toMatchObject({
      indentStartPt: 36,
      indentEndPt: 12,
      textIndentPt: -18,
      widthPt: 18,
    });
  });

  it("does not overlay the level's value when the paragraph specified its own", () => {
    const own = { ...numbered, indentStartPt: 35.4, textIndentPt: -18 };
    const [marker] = paragraphMarkers(listDoc(own), numbering);
    expect(marker).toMatchObject({
      indentStartPt: null,
      indentEndPt: 12,
      textIndentPt: null,
      widthPt: 18,
    });
  });

  it("takes only the hanging indent from the level when the paragraph specified just the left margin", () => {
    const own = { ...numbered, indentStartPt: 50 };
    const [marker] = paragraphMarkers(listDoc(own), numbering);
    expect(marker).toMatchObject({ indentStartPt: null, textIndentPt: -18 });
  });

  it("that indentation actually applies to the paragraph rendered on screen", () => {
    const doc = listDoc(numbered, NUMBERED_PPR);
    const mount = document.createElement("div");
    document.body.appendChild(mount);
    const view = createEditorView({
      mount,
      state: createEditorState(doc, { document: withLevels }),
      onStateChange: () => {},
    });
    mounted.view = view;

    const paragraph = mount.querySelector("p.docx-editor-p[data-marker]");
    if (!(paragraph instanceof HTMLElement)) throw new Error("no paragraph");
    expect(paragraph.style.marginInlineStart).toBe("36pt");
    expect(paragraph.style.marginInlineEnd).toBe("12pt");
    expect(paragraph.style.textIndent).toBe("-18pt");
  });

  it("the overlaid indentation does not remain in the document model", () => {
    const doc = listDoc(numbered, NUMBERED_PPR);
    const state = createEditorState(doc, { document: withLevels });
    const format = toParagraphFormat(state.doc.child(0).attrs.format);
    expect(format?.numbering).toEqual({ numId: 1, ilvl: 0 });
    expect(format?.indentStartPt).toBeUndefined();
    expect(format?.indentEndPt).toBeUndefined();
    expect(format?.textIndentPt).toBeUndefined();
  });
});

describe("what the level asks for around its number", () => {
  const styleOf = (properties: string) =>
    drawnParagraph(oneLevel(properties)).style;

  it("a level asking for a tab keeps the width its number sits in", () => {
    const style = styleOf(
      '<w:suff w:val="tab"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
    );
    expect(style.getPropertyValue(editorCssVariables.markerWidth)).toBe("18pt");
    expect(style.getPropertyValue(editorCssVariables.markerGap)).toBe("");
  });

  it("a level asking for a space keeps no width and no gap at all", () => {
    const style = styleOf(
      '<w:suff w:val="space"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
    );
    expect(style.getPropertyValue(editorCssVariables.markerWidth)).toBe("0");
    expect(style.getPropertyValue(editorCssVariables.markerGap)).toBe("0");
  });

  it("draws the space itself as part of the number", () => {
    const paragraph = drawnParagraph(oneLevel('<w:suff w:val="space"/>'));
    expect(paragraph.getAttribute(editorAttributes.listMarker)).toBe("1. ");
  });

  it("a level justifying its number to the end draws it against the text", () => {
    expect(
      styleOf('<w:lvlJc w:val="end"/>').getPropertyValue(
        editorCssVariables.markerAlign
      )
    ).toBe("right");
  });

  it("a level justifying it to the start leaves the rule's own alignment standing", () => {
    expect(
      styleOf('<w:lvlJc w:val="start"/>').getPropertyValue(
        editorCssVariables.markerAlign
      )
    ).toBe("");
  });
});

/**
 * A level dresses its own number and nothing else (§17.9.24), so what it writes has to reach the
 * marker without reaching the text of the paragraph the marker stands in front of.
 */
describe("the character formatting a level puts on its number", () => {
  /** A list of one level that draws its number bold, red, and larger than the text */
  const DRESSED_LEVEL =
    `<w:numbering ${W_NS}><w:abstractNum w:abstractNumId="0">` +
    '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
    '<w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="32"/></w:rPr>' +
    "</w:lvl></w:abstractNum>" +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

  const LIST_BODY = `<w:p>${NUMBERED_PPR}<w:r><w:t>Item</w:t></w:r></w:p>`;

  const opened = () => importDocx(makeNumberedDocx(LIST_BODY, DRESSED_LEVEL));

  /** The paragraph as the editor draws it, opened from a package rather than built by hand */
  function drawn(doc: PMNode, session: SessionStore): HTMLElement {
    const mount = document.createElement("div");
    document.body.appendChild(mount);
    mounted.view = createEditorView({
      mount,
      state: editorStateForSession({ doc, session }),
      onStateChange: () => {},
    });
    const paragraph = mount.querySelector("p.docx-editor-p[data-marker]");
    if (!(paragraph instanceof HTMLElement)) throw new Error("no paragraph");
    return paragraph;
  }

  it("the marker wears the level's rPr", () => {
    const { doc, session } = opened();
    const style = drawn(doc, session).style;

    expect(style.getPropertyValue(editorCssVariables.markerFontWeight)).toBe(
      "bold"
    );
    expect(style.getPropertyValue(editorCssVariables.markerColor)).toBe(
      "#FF0000"
    );
    expect(style.getPropertyValue(editorCssVariables.markerFontSize)).toBe(
      "16pt"
    );
  });

  it("the text of the paragraph the marker stands in front of does not", () => {
    const { doc } = opened();
    const mark = doc.child(0).child(0).marks[0];

    expect(toRunFormat(mark?.attrs.format)).toBeNull();
    expect(toParagraphFormat(doc.child(0).attrs.format)).toEqual({
      numbering: { numId: 1, ilvl: 0 },
    });
  });

  it("a level that dresses nothing leaves the number drawn like the text", () => {
    const style = drawnParagraph(oneLevel("")).style;

    expect(style.getPropertyValue(editorCssVariables.markerFontWeight)).toBe(
      ""
    );
    expect(style.getPropertyValue(editorCssVariables.markerColor)).toBe("");
  });

  it("a level switching a property off draws the number with it off", () => {
    const style = drawnParagraph(
      oneLevel('<w:rPr><w:b w:val="0"/><w:i w:val="0"/></w:rPr>')
    ).style;

    expect(style.getPropertyValue(editorCssVariables.markerFontWeight)).toBe(
      "normal"
    );
    expect(style.getPropertyValue(editorCssVariables.markerFontStyle)).toBe(
      "normal"
    );
  });
});

/**
 * The variables the decoration writes are read by the pseudo-element that draws the number, which
 * is the only thing that puts them on screen. A rule that stops reading one of them would leave
 * the decoration writing a value nothing draws.
 */
describe("the marker rule in editor.css", () => {
  const css = readFileSync(
    join(dirname(fileURLToPath(import.meta.url)), "../../styles/editor.css"),
    "utf8"
  );
  const rule = /\[data-marker\]::before\s*\{([^}]*)\}/.exec(css)?.[1] ?? "";

  it("draws the number from the attribute the decoration sets", () => {
    expect(rule).toContain(`attr(${editorAttributes.listMarker})`);
  });

  it.each([
    editorCssVariables.markerWidth,
    editorCssVariables.markerGap,
    editorCssVariables.markerAlign,
    editorCssVariables.markerFontWeight,
    editorCssVariables.markerFontStyle,
    editorCssVariables.markerColor,
    editorCssVariables.markerFontSize,
    editorCssVariables.markerFontFamily,
  ])("reads %s", (name) => {
    expect(rule).toContain(`var(${name}`);
  });
});

/** The span the first table occupies in the document */ /** The span the first table occupies in the document */
function tableRange(doc: PMNode): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  doc.forEach((block, offset) => {
    if (!range && block.type.name === "table") {
      range = { from: offset, to: offset + block.nodeSize };
    }
  });
  if (!range) throw new Error("no table");
  return range;
}

/**
 * Lists inside table cells.
 *
 * With the previous vendor, the first line of the paragraphs here appeared clipped.
 * Because we compute the numbers by walking the document and leave the drawing to the DOM,
 * that clipping is structurally impossible.
 */
describe("list numbers inside a table cell", () => {
  it("paragraphs inside a cell get numbers too", () => {
    const { doc, numbering } = openWithNumbering(LIST_FIXTURE);
    const { from, to } = tableRange(doc);
    const inCell = paragraphMarkers(doc, numbering).filter(
      (marker) => marker.from > from && marker.to < to
    );
    expect(inCell.map((marker) => marker.text)).toEqual(["1)", "2)", "3)"]);
    // The slot the number sits in comes from the paragraph's hanging indent (hanging 360 = 18pt)
    expect(inCell.map((marker) => marker.widthPt)).toEqual([18, 18, 18]);
  });

  it("counts the numbers inside and outside cells in the order they appear in the document", () => {
    const { doc, numbering } = openWithNumbering(LIST_FIXTURE);
    const markers = paragraphMarkers(doc, numbering);
    const sorted = [...markers].sort((a, b) => a.from - b.from);
    expect(markers).toEqual(sorted);
  });

  it("a numbered paragraph inside a cell renders whole without losing any text", () => {
    const { doc, session } = openWithNumbering(LIST_FIXTURE);
    const mount = document.createElement("div");
    document.body.appendChild(mount);
    const view = createEditorView({
      mount,
      state: editorStateForSession({ doc, session }),
      onStateChange: () => {},
    });
    mounted.view = view;

    const numbered = Array.from(
      mount.querySelectorAll("td .docx-editor-p[data-marker]")
    );
    expect(numbered.map((p) => p.getAttribute("data-marker"))).toEqual([
      "1)",
      "2)",
      "3)",
    ]);

    // The text rendered on screen does not differ from the document model's text by a single character
    const { from, to } = tableRange(doc);
    const modelTexts: string[] = [];
    doc.nodesBetween(from, to, (node) => {
      if (
        node.type.name === "paragraph" &&
        node.attrs.format?.numbering !== undefined
      ) {
        modelTexts.push(node.textContent);
      }
      return true;
    });
    expect(numbered.map((p) => p.textContent)).toEqual(modelTexts);

    // The numbers are decorations, so they never mix into the document text
    for (const text of modelTexts) expect(text.startsWith("1)")).toBe(false);
  });
});
