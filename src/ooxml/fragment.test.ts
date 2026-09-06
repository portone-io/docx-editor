// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { ANY_ELEMENT, ATTRIBUTES, acceptRawXml, ELEMENT } from "./fragment";
import { R_NS } from "./xml";

const SDT = {
  kind: "openTag",
  name: "sdt",
  closedBy: "<w:sdtContent/></w:sdt>",
  head: ["sdtPr", "sdtEndPr"],
} as const;
const HYPERLINK = {
  kind: "openTag",
  name: "hyperlink",
  closedBy: "</w:hyperlink>",
  head: [],
} as const;

describe("an element shape", () => {
  it("accepts a single well-formed w: element of the expected name", () => {
    const pPr = '<w:pPr><w:jc w:val="center"/></w:pPr>';

    expect(acceptRawXml(ELEMENT("pPr"), pPr)).toBe(pPr);
  });

  it("accepts any of the names it was given", () => {
    const shape = ELEMENT("footnoteReference", "endnoteReference");

    expect(acceptRawXml(shape, '<w:endnoteReference w:id="2"/>')).toBe(
      '<w:endnoteReference w:id="2"/>'
    );
  });

  it("refuses an element going by another name", () => {
    expect(acceptRawXml(ELEMENT("pPr"), "<w:rPr><w:b/></w:rPr>")).toBe(false);
  });

  it("refuses an element of another namespace under a named shape", () => {
    expect(acceptRawXml(ELEMENT("pPr"), "<m:pPr/>")).toBe(false);
  });

  it("accepts a single element of another namespace for an any shape", () => {
    expect(acceptRawXml(ANY_ELEMENT, "<m:oMathPara/>")).toBe("<m:oMathPara/>");
    expect(
      acceptRawXml(ANY_ELEMENT, '<w15:commentEx w15:paraId="0A" w15:done="0"/>')
    ).toBe('<w15:commentEx w15:paraId="0A" w15:done="0"/>');
  });

  it("refuses a fragment that closes its parent", () => {
    const smuggled =
      '</w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
      "<w:r><w:t>smuggled</w:t></w:r>";

    expect(acceptRawXml(ELEMENT("pPr"), smuggled)).toBe(false);
  });

  it("refuses a fragment with a sibling after the element", () => {
    expect(
      acceptRawXml(ELEMENT("rPr"), "<w:rPr><w:b/></w:rPr><w:t>smuggled</w:t>")
    ).toBe(false);
  });

  it("refuses text outside the element", () => {
    expect(acceptRawXml(ELEMENT("rPr"), "<w:rPr/>smuggled")).toBe(false);
    expect(acceptRawXml(ANY_ELEMENT, 'smuggled<w:bookmarkEnd w:id="1"/>')).toBe(
      false
    );
  });

  it("refuses a fragment that never closes", () => {
    expect(acceptRawXml(ELEMENT("pPr"), "<w:pPr>")).toBe(false);
  });

  it("refuses a fragment holding nothing at all", () => {
    expect(acceptRawXml(ANY_ELEMENT, "")).toBe(false);
  });

  it("refuses text standing alone where an element belongs", () => {
    expect(acceptRawXml(ANY_ELEMENT, "smuggled")).toBe(false);
  });

  it("carries an absent attr through untouched", () => {
    expect(acceptRawXml(ELEMENT("pPr"), null)).toBeNull();
    expect(acceptRawXml(ATTRIBUTES, null)).toBeNull();
  });
});

describe("an attributes shape", () => {
  it("accepts an attribute list and refuses one that opens a child", () => {
    const attrs = 'w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3"';

    expect(acceptRawXml(ATTRIBUTES, attrs)).toBe(attrs);
    expect(
      acceptRawXml(ATTRIBUTES, 'w:rsidR="00A"><w:r><w:t>smuggled</w:t></w:r')
    ).toBe(false);
  });

  it("declares the prefix of the first attribute", () => {
    const paraId = 'w14:paraId="1A2B3C4D"';

    expect(acceptRawXml(ATTRIBUTES, paraId)).toBe(paraId);
  });

  it("refuses a list that closes the tag it was written into", () => {
    expect(acceptRawXml(ATTRIBUTES, 'w:rsidR="00A"/><w:p x="')).toBe(false);
  });

  it("refuses an unbalanced quote", () => {
    expect(acceptRawXml(ATTRIBUTES, 'w:rsidR="00A')).toBe(false);
  });

  it("refuses a list that closes its own tag and opens a sibling", () => {
    expect(acceptRawXml(ATTRIBUTES, 'w:rsidR="00A"></y><y w:rsidR="00B"')).toBe(
      false
    );
  });
});

/**
 * A producer may declare a namespace on the element that uses it rather than on the root of the
 * part, so a fragment arrives carrying declarations of its own. One that agrees with what the
 * prefix already means travels; one that binds a prefix the writer reads under is turned down.
 */
describe("a namespace the fragment declares", () => {
  it("keeps an attribute list that declares the prefix it uses", () => {
    const paraId = 'xmlns:w14="u" w14:paraId="1"';

    expect(acceptRawXml(ATTRIBUTES, paraId)).toBe(paraId);
  });

  it("keeps an element that declares the prefix it uses", () => {
    const pPr = '<w:pPr xmlns:w14="u"><w14:conflictMode w:val="1"/></w:pPr>';

    expect(acceptRawXml(ELEMENT("pPr"), pPr)).toBe(pPr);
  });

  it("refuses a list binding the relationship prefix somewhere else", () => {
    expect(acceptRawXml(ATTRIBUTES, 'xmlns:r="urn:evil"')).toBe(false);
  });

  it("refuses a list binding the wordprocessing prefix somewhere else", () => {
    expect(acceptRawXml(ATTRIBUTES, 'xmlns:w="urn:evil"')).toBe(false);
  });

  it("refuses an element that rebinds a reserved prefix inside itself", () => {
    expect(
      acceptRawXml(ELEMENT("pPr"), '<w:pPr><w:jc xmlns:r="urn:evil"/></w:pPr>')
    ).toBe(false);
  });

  it("refuses an opening tag that rebinds a reserved prefix", () => {
    expect(acceptRawXml(HYPERLINK, '<w:hyperlink xmlns:r="urn:evil">')).toBe(
      false
    );
  });

  it("keeps a declaration that agrees with the prefix it binds", () => {
    const bound = `xmlns:r="${R_NS}"`;

    expect(acceptRawXml(ATTRIBUTES, bound)).toBe(bound);
  });
});

describe("an open tag shape", () => {
  it("accepts an sdt opening tag closed by sdtContent", () => {
    const prefix =
      '<w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

    expect(acceptRawXml(SDT, prefix)).toBe(prefix);
  });

  it("accepts a hyperlink opening tag closed by its own end tag", () => {
    const prefix = '<w:hyperlink r:id="rId7" w:history="1">';

    expect(acceptRawXml(HYPERLINK, prefix)).toBe(prefix);
  });

  it("refuses an opening tag that closes itself, leaving the closing text stranded", () => {
    expect(acceptRawXml(HYPERLINK, "<w:hyperlink/>")).toBe(false);
  });

  it("refuses a prefix that opens a second element beside the one it names", () => {
    expect(
      acceptRawXml(SDT, "<w:sdt><w:sdtPr/></w:sdt><w:sdt><w:sdtPr/>")
    ).toBe(false);
  });

  it("refuses an opening tag of another element", () => {
    expect(acceptRawXml(SDT, "<w:tbl>")).toBe(false);
  });

  it("accepts the properties a control carries ahead of its content", () => {
    const prefix =
      '<w:sdt w:foo="1"><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtEndPr/>';

    expect(acceptRawXml(SDT, prefix)).toBe(prefix);
  });

  it("refuses a link opening tag carrying a run of its own", () => {
    expect(
      acceptRawXml(
        HYPERLINK,
        '<w:hyperlink r:id="rId7"><w:r><w:t>smuggled</w:t></w:r>'
      )
    ).toBe(false);
  });

  it("refuses a control opening tag that brings its own content slot", () => {
    expect(
      acceptRawXml(
        SDT,
        "<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>smuggled</w:t></w:r></w:sdtContent>"
      )
    ).toBe(false);
  });

  it("refuses an opening tag showing text where its content goes", () => {
    expect(acceptRawXml(HYPERLINK, "<w:hyperlink>smuggled")).toBe(false);
  });
});
