// @vitest-environment jsdom
import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { afterEach, describe, expect, it } from "vitest";
import { bytesEqual, readFixture } from "../../__testing__/docx";
import { exportDocx } from "../../docx/exportDocx";
import { importDocx } from "../../docx/importDocx";
import { toParagraphFormat } from "../../model/format";
import { type Numbering, parseNumbering } from "../../numbering/parseNumbering";
import { docxSchema } from "../../schema";
import { createEditorState, createEditorView } from "../createEditor";
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
      const { bytes, doc, session, numbering } = openWithNumbering(name);
      const out = exportDocx(
        createEditorState(doc, { numbering }).doc,
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
    const { doc, numbering } = openWithNumbering(LIST_FIXTURE);
    const withMarkers = createEditorState(doc, { numbering });
    const withoutMarkers = createEditorState(doc);
    expect(withMarkers.doc.textContent).toBe(withoutMarkers.doc.textContent);
  });

  it("recounts the numbers that follow when a paragraph disappears", () => {
    const { doc, numbering } = openWithNumbering(LIST_FIXTURE);
    const state = createEditorState(doc, { numbering });
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

/**
 * When a paragraph has not written down a `w:ind`, Word uses the indentation the list level
 * specifies. Without overlaying that value on screen, the number would sit at the paragraph's
 * left edge and the whole body text would be pushed over by the number's width.
 */
describe("the indentation the level specifies", () => {
  const W_NS =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

  const numbering: Numbering = parseNumbering(
    `<w:numbering ${W_NS}><w:abstractNum w:abstractNumId="0">` +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
      '<w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:right="240" w:hanging="360"/></w:pPr>' +
      "</w:lvl></w:abstractNum>" +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
  );

  function listDoc(format: Record<string, unknown>): PMNode {
    return docxSchema.nodes.doc.create(null, [
      docxSchema.nodes.paragraph.create(
        { srcId: 0, pAttrs: null, pPr: null, format },
        docxSchema.text("Item")
      ),
    ]);
  }

  const numbered = { numbering: { numId: 1, ilvl: 0 } };

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
    const doc = listDoc(numbered);
    const mount = document.createElement("div");
    document.body.appendChild(mount);
    const view = createEditorView({
      mount,
      state: createEditorState(doc, { numbering }),
      defaults: { fontSizePt: null, fontFamily: null, lineSpacing: null },
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
    const doc = listDoc(numbered);
    const state = createEditorState(doc, { numbering });
    expect(toParagraphFormat(state.doc.child(0).attrs.format)).toEqual({
      numbering: { numId: 1, ilvl: 0 },
    });
  });
});

/** The span the first table occupies in the document */
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
    const { doc, session, numbering } = openWithNumbering(LIST_FIXTURE);
    const mount = document.createElement("div");
    document.body.appendChild(mount);
    const view = createEditorView({
      mount,
      state: createEditorState(doc, { numbering }),
      defaults: session.defaults,
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
