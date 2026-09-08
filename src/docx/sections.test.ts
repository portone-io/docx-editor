// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import {
  LETTER_GEOMETRY as LETTER,
  LETTER_SECT_PR,
  makeDocx,
} from "../__testing__/docx";
import { childByLocalName, parseXml } from "../ooxml/xml";
import { importDocx } from "./importDocx";
import {
  firstSectPrElement,
  parseSectionProperties,
  readSectionProperties,
  withoutSectionBreak,
  withSectionBreak,
} from "./sections";

/** A4 landscape, the paper a second section is given here so the two cannot be confused */
const A4_LANDSCAPE_SECT_PR =
  "<w:sectPr>" +
  '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
  '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/>' +
  '<w:pgNumType w:start="5"/>' +
  '<w:type w:val="nextPage"/>' +
  "</w:sectPr>";

/**
 * Two sections: a Letter one closed by the first paragraph, and an A4 landscape one closed by the
 * body. This is the shape §17.6.17 and §17.6.18 lay down together.
 */
const TWO_SECTIONS =
  `<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>` +
  "<w:r><w:t>first section</w:t></w:r></w:p>" +
  "<w:p><w:r><w:t>second section</w:t></w:r></w:p>" +
  A4_LANDSCAPE_SECT_PR;

/** The first section of a document, read the way the import reads it */
function firstSectionOf(body: string) {
  const documentXml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body}</w:body></w:document>`;
  const element = childByLocalName(
    parseXml(documentXml).documentElement,
    "body"
  );
  if (!element) throw new Error("the document has no body");
  const sectPr = firstSectPrElement(element);
  if (!sectPr) throw new Error("the document names no section");
  return readSectionProperties(sectPr);
}

describe("the section properties one w:sectPr lays down", () => {
  it("reads the section a paragraph closes ahead of the one closing the body", () => {
    expect(firstSectionOf(TWO_SECTIONS).geometry).toEqual(LETTER);
    expect(firstSectionOf(A4_LANDSCAPE_SECT_PR).pageNumberStart).toBe(5);
    expect(firstSectionOf(A4_LANDSCAPE_SECT_PR).type).toBe("nextPage");
  });

  it("reads the header and footer stories the section selects, one per variant", () => {
    const props = firstSectionOf(
      "<w:sectPr>" +
        '<w:headerReference w:type="first" r:id="rId4"/>' +
        '<w:headerReference w:type="default" r:id="rId5"/>' +
        '<w:footerReference r:id="rId6"/>' +
        "<w:titlePg/>" +
        "</w:sectPr>"
    );

    expect(props.headerRefs).toEqual({
      default: "rId5",
      first: "rId4",
      even: null,
    });
    // A reference naming no variant names the default one (§17.6.12)
    expect(props.footerRefs.default).toBe("rId6");
    expect(props.titlePg).toBe(true);
  });

  it("reads a page-number start no counter could reach as none at all", () => {
    // `1` with 309 zeroes after it is `Infinity` once read, and a page counter starting there
    // would draw a page number nothing could follow
    const props = firstSectionOf(
      '<w:sectPr><w:pgNumType w:start="1e309"/><w:type w:val="sideways"/></w:sectPr>'
    );

    expect(props.pageNumberStart).toBeNull();
    // §17.6.22 names five kinds, and a section claiming another names none
    expect(props.type).toBeNull();
  });

  it("keeps the text a section was written as when it is read out of a fragment", () => {
    const props = parseSectionProperties(A4_LANDSCAPE_SECT_PR);

    expect(props?.xml).toBe(A4_LANDSCAPE_SECT_PR);
    expect(props?.geometry.widthTwips).toBe(16838);
    expect(parseSectionProperties("<w:pPr/>")).toBeNull();
    expect(parseSectionProperties("not xml at all")).toBeNull();
  });

  it("the session's paper is the first section's", () => {
    const { session } = importDocx(makeDocx(TWO_SECTIONS));

    expect(session.geometry).toEqual(LETTER);
    expect(session.geometry).toEqual(firstSectionOf(TWO_SECTIONS).geometry);
  });

  it("moves a section break between the halves a split leaves", () => {
    const pPr = `<w:pPr><w:jc w:val="center"/>${LETTER_SECT_PR}</w:pPr>`;

    expect(withoutSectionBreak(pPr)).toBe(
      '<w:pPr><w:jc w:val="center"/></w:pPr>'
    );
    // Properties holding nothing else collapse rather than going out empty
    expect(withoutSectionBreak(`<w:pPr>${LETTER_SECT_PR}</w:pPr>`)).toBeNull();
    // The break goes back into the spot CT_PPr gives it, after the direct formatting
    expect(
      withSectionBreak('<w:pPr><w:jc w:val="center"/></w:pPr>', LETTER_SECT_PR)
    ).toBe(pPr);
    expect(withSectionBreak(null, LETTER_SECT_PR)).toBe(
      `<w:pPr>${LETTER_SECT_PR}</w:pPr>`
    );
  });
});
