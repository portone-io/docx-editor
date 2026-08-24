// @vitest-environment jsdom
import { unzipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  decode,
  fixtureNames,
  LETTER_GEOMETRY as LETTER,
  LETTER_FIXTURE,
  LETTER_SECT_PR,
  makeDocx,
  readFixture,
} from "../__testing__/docx";
import { childByLocalName, parseXml } from "../ooxml/xml";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import {
  A4_BODY_WIDTH,
  A4_PORTRAIT,
  bodyHeightTwips,
  bodyWidth,
  bodyWidthTwips,
  type PageGeometry,
  readBodyGeometry,
} from "./pageGeometry";

function bodyWith(sectPr: string): string {
  return `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/>${sectPr}</w:body></w:document>`;
}

/** The reading as the importer does it: off the body of a document already parsed */
function geometryOf(documentXml: string): PageGeometry {
  const body = childByLocalName(parseXml(documentXml).documentElement, "body");
  if (!body) throw new Error("the document has no body");
  return readBodyGeometry(body);
}

describe("the page geometry a document lays down", () => {
  it("reads the paper and the margins a section names", () => {
    expect(geometryOf(bodyWith(LETTER_SECT_PR))).toEqual(LETTER);
  });

  it("draws a document that names no section on A4", () => {
    expect(geometryOf(bodyWith(""))).toEqual(A4_PORTRAIT);
  });

  it("draws a document whose section names no size on A4", () => {
    const geometry = geometryOf(bodyWith("<w:sectPr/>"));
    expect(geometry).toEqual(A4_PORTRAIT);
  });

  it("takes the margins a section names even where it names no size", () => {
    const geometry = geometryOf(
      bodyWith('<w:sectPr><w:pgMar w:left="720" w:right="720"/></w:sectPr>')
    );
    expect(geometry.widthTwips).toBe(A4_PORTRAIT.widthTwips);
    expect(geometry.marginLeftTwips).toBe(720);
    expect(geometry.marginRightTwips).toBe(720);
    // The ones it left unsaid stay as A4 wrote them
    expect(geometry.marginTopTwips).toBe(A4_PORTRAIT.marginTopTwips);
  });

  it("takes a landscape paper as the section wrote it, already turned", () => {
    // Word writes the width and the height swapped, so w:orient is not applied a second time
    const geometry = geometryOf(
      bodyWith(
        '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/></w:sectPr>'
      )
    );
    expect(geometry.widthTwips).toBe(15840);
    expect(geometry.heightTwips).toBe(12240);
    // The body is wider than it is tall, which is what landscape means
    expect(bodyWidthTwips(geometry)).toBeGreaterThan(bodyHeightTwips(geometry));
  });

  it("draws a paper too small to hold a body on A4", () => {
    const geometry = geometryOf(
      bodyWith('<w:sectPr><w:pgSz w:w="200" w:h="200"/></w:sectPr>')
    );
    expect(geometry).toEqual(A4_PORTRAIT);
  });

  it("draws a paper its margins leave no room on on A4", () => {
    const geometry = geometryOf(
      bodyWith(
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="9000" w:right="9000"/></w:sectPr>'
      )
    );
    expect(geometry).toEqual(A4_PORTRAIT);
  });

  it("takes a negative margin as none, the way a gutter is not drawn", () => {
    const geometry = geometryOf(
      bodyWith(
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="-720" w:right="1440"/></w:sectPr>'
      )
    );
    expect(geometry.marginLeftTwips).toBe(0);
    expect(geometry.marginRightTwips).toBe(1440);
  });

  it("takes the first section where a document holds several", () => {
    const twoSections = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr></w:pPr></w:p><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
    expect(geometryOf(twoSections)).toEqual(LETTER);
  });

  it("gives the body the paper less its margins", () => {
    expect(bodyWidthTwips(LETTER)).toBe(12240 - 1440 - 1440);
    expect(bodyHeightTwips(LETTER)).toBe(15840 - 1440 - 1440);
  });

  it("hands out the body width in the twips a table wants and the pixels the sheet wants", () => {
    // 9360 twips is 6.5in, which CSS draws as 6.5 * 96 pixels
    expect(bodyWidth(LETTER)).toEqual({ twips: 9360, px: 624 });
  });

  it("gives A4 the body width every document used to be fitted to", () => {
    expect(A4_BODY_WIDTH.twips).toBe(9412);
    expect(A4_BODY_WIDTH).toEqual(bodyWidth(A4_PORTRAIT));
    // The paper that is not A4 is fitted to its own width, not to that one
    expect(bodyWidth(LETTER).twips).not.toBe(A4_BODY_WIDTH.twips);
  });
});

function exportedDocumentXml(bytes: Uint8Array): string {
  const { doc, session } = importDocx(bytes);
  return decode(unzipSync(exportDocx(doc, session))["word/document.xml"]);
}

/**
 * A document of two sections: the first names Letter and carries its `w:sectPr` on the last
 * paragraph of the section, the last names A4 at the end of the body, as Word writes them.
 */
const TWO_SECTIONS =
  `<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr><w:r><w:t xml:space="preserve">first</w:t></w:r></w:p>` +
  '<w:p><w:r><w:t xml:space="preserve">second</w:t></w:r></w:p>' +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

describe("a document written in several sections", () => {
  it("draws the whole document on the first section's paper", () => {
    const { doc, session } = importDocx(makeDocx(TWO_SECTIONS));

    expect(session.geometry).toEqual(LETTER);
    // One paper for the document, so the A4 the last section names reaches nothing on screen
    expect(session.geometry.heightTwips).not.toBe(16838);
    expect(doc.childCount).toBe(2);
  });

  it("carries both sections back out untouched", () => {
    const exported = exportedDocumentXml(makeDocx(TWO_SECTIONS));

    expect(exported).toContain(LETTER_SECT_PR);
    expect(exported).toContain(
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
    );
  });
});

/** The paper the whole document is drawn on, read once when the file is opened */
function fixtureGeometry(name: string): PageGeometry {
  return importDocx(readFixture(name)).session.geometry;
}

describe("the paper the fixtures are written on", () => {
  it("reads Letter off the one document written on Letter", () => {
    expect(fixtureGeometry(LETTER_FIXTURE)).toEqual(LETTER);
    expect(bodyWidth(fixtureGeometry(LETTER_FIXTURE)).twips).toBe(9360);
  });

  it("reads A4 with an inch of margin off the other three", () => {
    for (const name of fixtureNames.filter((one) => one !== LETTER_FIXTURE)) {
      const geometry = fixtureGeometry(name);
      expect(geometry.widthTwips).toBe(11906);
      expect(geometry.heightTwips).toBe(16838);
      // Their own margins, not the 2.2cm of the fallback: 11906 - 1440 * 2
      expect(bodyWidth(geometry).twips).toBe(9026);
      expect(bodyWidth(geometry).twips).not.toBe(A4_BODY_WIDTH.twips);
    }
  });

  it("never writes the numbers it read back into the document", () => {
    const exported = exportedDocumentXml(readFixture(LETTER_FIXTURE));

    // The section rides out as the bytes it came in as, attribute order and all
    expect(exported).toContain(
      '<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>'
    );
    expect(exported).toContain(
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    );
  });
});
