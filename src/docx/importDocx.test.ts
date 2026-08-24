// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  importErrorCode,
  makeDocx,
  makeStyledDocx,
  readFixture,
} from "../__testing__/docx";
import { toParagraphFormat, toRunFormat } from "../model/format";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";

const encode = (text: string) => new TextEncoder().encode(text);

/** A zip holding exactly the parts given, so a docx can be broken one part at a time */
function makePackage(parts: Record<string, string>): Uint8Array {
  return zipSync(
    Object.fromEntries(
      Object.entries(parts).map(([path, text]) => [path, encode(text)])
    )
  );
}

const PACKAGE_RELS =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Target="word/document.xml" ' +
  'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>' +
  "</Relationships>";

const W_NS_DECL =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe("refusing to open", () => {
  it("explicitly rejects bytes that are not a docx", () => {
    const notADocx = encode("this is not a docx");
    expect(importErrorCode(() => importDocx(notADocx))).toBe("not-a-docx");
  });

  it("explicitly rejects a broken file too", () => {
    const truncated = makeDocx("<w:p/>").slice(0, 40);
    expect(importErrorCode(() => importDocx(truncated))).toBe("not-a-docx");
  });

  it("rejects a package with no relationship pointing at the main part", () => {
    const withoutRels = makePackage({ "word/document.xml": "<w:document/>" });
    expect(importErrorCode(() => importDocx(withoutRels))).toBe("missing-part");
  });

  it("rejects a package whose relationship points at a main part that is not there", () => {
    const withoutMainPart = makePackage({ "_rels/.rels": PACKAGE_RELS });
    expect(importErrorCode(() => importDocx(withoutMainPart))).toBe(
      "missing-part"
    );
  });

  it("rejects a document with no body", () => {
    const withoutBody = makePackage({
      "_rels/.rels": PACKAGE_RELS,
      "word/document.xml": `<w:document ${W_NS_DECL}/>`,
    });
    expect(importErrorCode(() => importDocx(withoutBody))).toBe("missing-body");
  });

  it("rejects malformed XML", () => {
    const malformed = makePackage({
      "_rels/.rels": PACKAGE_RELS,
      "word/document.xml": `<w:document ${W_NS_DECL}><w:body><w:p></w:body></w:document>`,
    });
    expect(importErrorCode(() => importDocx(malformed))).toBe("malformed-xml");
  });

  /**
   * The parser expands the entities a DTD declares, so a document carrying one would show
   * a figure that stands nowhere in the part the export writes back
   */
  it("rejects a document whose main part declares a DTD", () => {
    const withDtd = makePackage({
      "_rels/.rels": PACKAGE_RELS,
      "word/document.xml":
        '<!DOCTYPE w:document [<!ENTITY fee "one hundred">]>' +
        `<w:document ${W_NS_DECL}><w:body><w:p><w:r>` +
        "<w:t>&fee;</w:t></w:r></w:p></w:body></w:document>",
    });
    expect(importErrorCode(() => importDocx(withDtd))).toBe("malformed-xml");
  });

  it("rejects a document holding an XML node we cannot write back out", () => {
    // A processing instruction inside a fragment we preserve is markup we have no way to write back out
    const withProcessingInstruction = makeDocx(
      "<w:p><w:pPr><?ignore me?></w:pPr></w:p>"
    );
    expect(importErrorCode(() => importDocx(withProcessingInstruction))).toBe(
      "unsupported-content"
    );
  });
});

describe("document default formatting", () => {
  it("the defaults are empty when there is no styles.xml", () => {
    const { session } = importDocx(makeDocx("<w:p/>"));
    expect(session.defaults).toEqual({
      fontSizePt: null,
      fontFamily: null,
      lineSpacing: null,
    });
  });

  it("carries the docDefaults of styles.xml into the session", () => {
    const { session } = importDocx(
      makeDocx(
        "<w:p/>",
        '<w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr>'
      )
    );
    expect(session.defaults).toEqual({
      fontSizePt: 10,
      fontFamily: '"Arial"',
      lineSpacing: null,
    });
  });

  it("applies pPrDefault values to paragraphs and retains the layer in the session", () => {
    const { doc, session } = importDocx(
      makeStyledDocx(
        "<w:p/>",
        "<w:docDefaults><w:pPrDefault><w:pPr>" +
          '<w:tabs><w:tab w:val="center" w:pos="1440"/></w:tabs>' +
          "</w:pPr></w:pPrDefault></w:docDefaults>"
      )
    );

    const expected = {
      tabStops: [{ positionPt: 72, align: "center" }],
    };
    expect(session.paragraphDefaults).toEqual(expected);
    expect(toParagraphFormat(doc.child(0).attrs.format)).toEqual(expected);
  });
});

describe("document tab interval", () => {
  it("uses the OOXML fallback when settings.xml does not declare an interval", () => {
    expect(importDocx(makeDocx("<w:p/>")).session.defaultTabStopPt).toBe(36);
  });

  it("follows the settings relationship and reads the declared interval", () => {
    const parts = unzipSync(makeDocx("<w:p/>"));
    parts["word/_rels/document.xml.rels"] = encode(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId7" Target="settings.xml" ' +
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"/>' +
        "</Relationships>"
    );
    parts["word/settings.xml"] = encode(
      `<w:settings ${W_NS_DECL}><w:defaultTabStop w:val="960"/></w:settings>`
    );

    expect(importDocx(zipSync(parts)).session.defaultTabStopPt).toBe(48);
  });
});

/** The run mark attached to the first piece of text */
function firstRunMark(paragraph: PMNode) {
  const mark = paragraph.child(0).marks[0];
  if (!mark) throw new Error("no run mark");
  return mark;
}

/**
 * The shape almost every table Word writes has: the body holds a `w:tblStyle` reference and the
 * lines themselves are defined in styles.xml. Nothing here draws a border of its own
 */
const GRID_TABLE_BODY =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>' +
  '<w:tblW w:type="auto" w:w="0"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">cell</w:t></w:r></w:p></w:tc></w:tr>' +
  "</w:tbl>";

const TABLE_GRID_XML =
  '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:tblPr/></w:style>' +
  '<w:style w:type="table" w:styleId="TableGrid"><w:basedOn w:val="TableNormal"/>' +
  "<w:tblPr><w:tblBorders>" +
  ["top", "left", "bottom", "right", "insideH", "insideV"]
    .map(
      (name) =>
        `<w:${name} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`
    )
    .join("") +
  "</w:tblBorders></w:tblPr></w:style>";

const GRID_LINE = "0.5pt solid #000000";

describe("resolving the style chain", () => {
  it("the styles are empty too when there is no styles.xml", () => {
    const { session } = importDocx(makeDocx("<w:p/>"));
    expect(session.styles.size).toBe(0);
  });

  it("layers the values of the style the paragraph points at underneath the display values", () => {
    // The title paragraph is the only one in the fixture that points at a style
    const { doc, session } = importDocx(readFixture("kitchen-sink.docx"));
    expect(session.styles.get("Heading1")?.run).toEqual({
      color: "#2E74B5",
      fontSizePt: 16,
    });

    const title = doc.child(0);
    const mark = firstRunMark(title);
    // The color the style gave survives, while for the size the value the paragraph wrote down wins
    expect(toRunFormat(mark.attrs.format)).toEqual({
      bold: true,
      fontSizePt: 18,
      color: "#2E74B5",
    });
    // The original formatting XML is left untouched
    expect(mark.attrs.rPr).toBe(
      '<w:rPr><w:b/><w:bCs/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
    );
  });

  it("a paragraph that points at no style is left as it is", () => {
    const { doc } = importDocx(readFixture("kitchen-sink.docx"));
    const second = doc.child(1);
    expect(toRunFormat(firstRunMark(second).attrs.format)).toEqual({
      fontSizePt: 11,
    });
  });

  it("a paragraph inside a table cell goes down the same path", () => {
    const bytes = makeStyledDocx(
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
        "<w:tr><w:tc>" +
        '<w:p><w:pPr><w:pStyle w:val="Cell"/></w:pPr>' +
        '<w:r><w:t xml:space="preserve">cell</w:t></w:r></w:p>' +
        "</w:tc></w:tr></w:tbl>",
      '<w:style w:styleId="Base"><w:pPr><w:jc w:val="center"/></w:pPr>' +
        '<w:rPr><w:sz w:val="30"/></w:rPr></w:style>' +
        '<w:style w:styleId="Cell"><w:basedOn w:val="Base"/>' +
        '<w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style>'
    );
    const { doc } = importDocx(bytes);
    const paragraph = doc.child(0).child(0).child(0).child(0);
    expect(toParagraphFormat(paragraph.attrs.format)).toEqual({
      align: "center",
    });
    expect(toRunFormat(firstRunMark(paragraph).attrs.format)).toEqual({
      fontSizePt: 15,
      highlight: "yellow",
    });
  });

  /**
   * OOXML applies the default paragraph style to every paragraph pointing at no style of its own,
   * so a document read without it is drawn differently from the way Word draws it
   */
  it("lays the default paragraph style under a paragraph that points at none", () => {
    const bytes = makeStyledDocx(
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
        '<w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">cell</w:t></w:r></w:p>' +
        "</w:tc></w:tr></w:tbl>" +
        '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>',
      '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
        '<w:name w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr>' +
        '<w:rPr><w:sz w:val="22"/></w:rPr></w:style>'
    );
    const { doc, session } = importDocx(bytes);
    expect(session.defaultParagraphStyleId).toBe("Normal");

    const body = doc.child(1);
    expect(toParagraphFormat(body.attrs.format)).toEqual({ align: "center" });
    // Baked onto the paragraph as well, so text typed into it is drawn in the style's formatting
    expect(toRunFormat(body.attrs.styleRun)).toEqual({ fontSizePt: 11 });
    expect(toRunFormat(firstRunMark(body).attrs.format)).toEqual({
      fontSizePt: 11,
    });

    // A paragraph inside a table cell wears it the same way
    const inCell = doc.child(0).child(0).child(0).child(0);
    expect(toParagraphFormat(inCell.attrs.format)).toEqual({ align: "center" });
    expect(toRunFormat(inCell.attrs.styleRun)).toEqual({ fontSizePt: 11 });
  });

  it("has no default paragraph style where the document marks none", () => {
    const { session } = importDocx(makeDocx("<w:p/>"));
    expect(session.defaultParagraphStyleId).toBeNull();
  });

  it("a table pointing at nothing but a table style still draws its grid", () => {
    const { doc } = importDocx(makeStyledDocx(GRID_TABLE_BODY, TABLE_GRID_XML));
    const table = doc.child(0);
    expect(table.type.name).toBe("table");
    expect(table.attrs.styleInside).toEqual({
      horizontal: GRID_LINE,
      vertical: GRID_LINE,
    });
    expect(table.child(0).child(0).attrs.format).toEqual({
      borderTop: GRID_LINE,
      borderBottom: GRID_LINE,
      borderLeft: GRID_LINE,
      borderRight: GRID_LINE,
    });
    expect(table.attrs.format).toMatchObject({ borderTop: GRID_LINE });
  });

  it("a table takes the cell margins its style chain lays down", () => {
    const margin = (name: string, twips: number) =>
      `<w:${name} w:w="${twips}" w:type="dxa"/>`;
    const { doc } = importDocx(
      makeStyledDocx(
        GRID_TABLE_BODY,
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:tblPr>' +
          "<w:tblCellMar>" +
          margin("top", 0) +
          margin("left", 108) +
          margin("bottom", 0) +
          margin("right", 108) +
          "</w:tblCellMar></w:tblPr></w:style>" +
          '<w:style w:type="table" w:styleId="TableGrid"><w:basedOn w:val="TableNormal"/>' +
          `<w:tblPr><w:tblCellMar>${margin("left", 200)}</w:tblCellMar>` +
          "</w:tblPr></w:style>"
      )
    );
    const table = doc.child(0);
    expect(table.attrs.styleCellMargins).toEqual({
      topPt: 0,
      rightPt: 5.4,
      bottomPt: 0,
      leftPt: 10,
    });
    // The table itself writes no margin, so every side comes from the chain above it
    expect(table.child(0).child(0).attrs.format).toEqual({
      paddingTopPt: 0,
      paddingRightPt: 5.4,
      paddingBottomPt: 0,
      paddingLeftPt: 10,
    });
  });

  it("the lines that came from the table style do not disturb the exported bytes", () => {
    const bytes = makeStyledDocx(GRID_TABLE_BODY, TABLE_GRID_XML);
    const { doc, session } = importDocx(bytes);
    const original = unzipSync(bytes);
    const exported = unzipSync(exportDocx(doc, session));
    expect(Object.keys(exported)).toEqual(Object.keys(original));
    for (const key of Object.keys(original)) {
      expect(bytesEqual(exported[key], original[key])).toBe(true);
    }
  });

  it("resolving styles does not disturb the exported bytes", () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
      '<w:r><w:t xml:space="preserve">heading</w:t></w:r></w:p>';
    const { doc, session } = importDocx(
      makeStyledDocx(
        body,
        '<w:style w:styleId="Heading1"><w:rPr><w:b/></w:rPr></w:style>'
      )
    );
    // Display values that came from a style do not disturb the node comparison, so the original fragment goes out unchanged
    expect(doc.child(0).eq(session.blocks[0].node)).toBe(true);
    expect(session.blocks[0].xml).toBe(body);
  });
});

/** The Japanese and Chinese fixture, the one document whose fonts come out of a theme */
const CJK_FIXTURE = "east-asian.docx";

/** The first paragraph whose text begins with this, so a test names what it reads */
function paragraphStartingWith(doc: PMNode, prefix: string): PMNode {
  let found: PMNode | null = null;
  doc.forEach((block) => {
    if (found === null && block.type.name === "paragraph") {
      if (block.textContent.startsWith(prefix)) found = block;
    }
  });
  if (found === null) throw new Error(`no paragraph starting with ${prefix}`);
  return found;
}

/** The font list the first run of one paragraph asks for */
function fontFamilyOf(doc: PMNode, prefix: string): string | undefined {
  const format = toRunFormat(
    firstRunMark(paragraphStartingWith(doc, prefix)).attrs.format
  );
  return format?.fontFamily;
}

describe("the fonts a theme names", () => {
  it("a run pointing at nothing but the theme is drawn in the theme fonts", () => {
    const { doc } = importDocx(readFixture(CJK_FIXTURE));
    // The body scheme for a paragraph, the heading scheme for a heading
    expect(fontFamilyOf(doc, "鞄から")).toBe('"Calibri","Yu Mincho"');
    expect(fontFamilyOf(doc, "Japanese")).toBe('"Cambria","Yu Gothic"');
  });

  it("a run that names its fonts keeps the name of every slot", () => {
    const { doc } = importDocx(readFixture(CJK_FIXTURE));
    expect(fontFamilyOf(doc, "電車は")).toBe('"Times New Roman","MS Mincho"');
    expect(fontFamilyOf(doc, "茶还是")).toBe('"Times New Roman","SimSun"');
  });

  it("the document default font is resolved out of the theme as well", () => {
    const { session } = importDocx(readFixture(CJK_FIXTURE));
    expect(session.defaults.fontFamily).toBe('"Calibri","Yu Mincho"');
  });

  it("resolving them writes the references back out, not the names", () => {
    const { doc, session } = importDocx(readFixture(CJK_FIXTURE));
    const exported = unzipSync(exportDocx(doc, session));
    // The references themselves are what the file keeps; only the screen sees a name
    expect(decode(exported[session.mainPartPath])).toContain(
      'w:eastAsiaTheme="minorEastAsia"'
    );
    expect(decode(exported[session.mainPartPath])).not.toContain("Yu Mincho");
  });
});
