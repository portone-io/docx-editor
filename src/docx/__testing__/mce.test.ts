// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { parseXml, W_NS } from "../../ooxml/xml";
import { withoutIgnorableMarkup } from "./mce";

const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const documentWith = (body: string, attributes = "") =>
  `<w:document xmlns:w="${W_NS}" xmlns:mc="${MC}" xmlns:x="urn:unknown" ${attributes}><w:body>${body}</w:body></w:document>`;
const process = (body: string, attributes = "") =>
  withoutIgnorableMarkup(documentWith(body, attributes));

function fallback(body: string, attributes = ""): string {
  return `<mc:AlternateContent ${attributes}><mc:Choice Requires="x"/><mc:Fallback>${body}</mc:Fallback></mc:AlternateContent>`;
}

describe("the MCE validation profile", () => {
  it("keeps understood markup even when declared ignorable", () => {
    const out = process('<w:p mc:Ignorable="w"><w:bad w:val="invalid"/></w:p>');
    expect(out).toContain('<w:bad w:val="invalid"/>');
  });

  it("selects the first understood choice, including a schema-imported namespace", () => {
    const out = process(
      '<w:p><mc:AlternateContent xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><mc:Choice Requires="x"><w:t>unknown</w:t></mc:Choice><mc:Choice Requires="w m"><w:bad/></mc:Choice><mc:Choice Requires="w"><w:t>later</w:t></mc:Choice><mc:Fallback><w:t>fallback</w:t></mc:Fallback></mc:AlternateContent></w:p>'
    );
    expect(out).toContain("<w:bad/>");
    expect(out).not.toMatch(/unknown<|later<|fallback<|AlternateContent/);
  });

  it("processes an element's own Ignorable declaration", () => {
    expect(process('<w:p><x:discard mc:Ignorable="x"/></w:p>')).not.toContain(
      "x:discard"
    );
  });

  it("inherits namespace URIs, not rebound prefix spellings", () => {
    const out = process(
      '<w:p><x:kept xmlns:x="urn:other"/><y:discard xmlns:y="urn:unknown"/></w:p>',
      'mc:Ignorable="x"'
    );
    expect(out).toContain("x:kept");
    expect(out).not.toContain("y:discard");
  });

  it("unions nested ignorable declarations", () => {
    expect(
      process(
        '<w:p xmlns:y="urn:second" mc:Ignorable="y"><x:one/><y:two/></w:p>',
        'mc:Ignorable="x"'
      )
    ).not.toMatch(/x:one|y:two/);
  });

  it.each(["x:wrap", "x:*"])(
    "unwraps ProcessContent %s without losing invalid WML",
    (name) => {
      const out = process(
        "<w:p><x:wrap><w:bad/></x:wrap></w:p>",
        `mc:Ignorable="x" mc:ProcessContent="${name}"`
      );
      expect(out).toContain("<w:bad/>");
      expect(out).not.toContain("x:wrap");
    }
  );

  it("resolves ProcessContent declared on the element itself", () => {
    expect(
      process(
        '<w:p><x:wrap mc:Ignorable="x" mc:ProcessContent="x:wrap"><w:bad/></x:wrap></w:p>'
      )
    ).toContain("<w:bad/>");
  });

  it.each(['mc:Ignorable="x"', 'xmlns:y="urn:other" mc:Ignorable="y"'])(
    "retains alternate-content scope %s until children are processed",
    (attributes) => {
      const prefix = attributes.includes("xmlns:y") ? "y" : "x";
      const out = process(
        `<w:p>${fallback(`<w:r ${prefix}:flag="1"/>`, attributes)}</w:p>`
      );
      expect(out).not.toContain(":flag");
    }
  );

  it("retains fallback declarations and nested alternate content", () => {
    const inner = fallback('<w:r x:flag="1"/>');
    const out = process(
      `<w:p><mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback mc:Ignorable="x">${inner}</mc:Fallback></mc:AlternateContent></w:p>`
    );
    expect(out).not.toMatch(/:flag|AlternateContent/);
  });

  it("retains fallback text so validation can reject it in element-only content", () => {
    expect(process(`<w:p>${fallback("invalid text<w:r/>")}</w:p>`)).toContain(
      "invalid text"
    );
  });

  it("discards an unselected alternative with no fallback", () => {
    expect(
      process(
        '<w:p><mc:AlternateContent><mc:Choice Requires="x"><w:bad/></mc:Choice></mc:AlternateContent></w:p>'
      )
    ).not.toContain("w:bad");
  });

  it.each([
    [
      '<w:p><x:wrap xml:space="preserve" mc:Ignorable="x" mc:ProcessContent="x:wrap"><w:r/></x:wrap></w:p>',
      "",
    ],
    [
      '<w:p><mc:AlternateContent unexpected="1"><mc:Choice Requires="w"/></mc:AlternateContent></w:p>',
      "",
    ],
    [
      '<w:p><mc:AlternateContent><mc:Choice Requires="w" x:flag="1"/></mc:AlternateContent></w:p>',
      "",
    ],
    [
      '<w:p><mc:AlternateContent><mc:Choice Requires="w" xml:space="preserve"/></mc:AlternateContent></w:p>',
      "",
    ],
    ["<w:p/>", 'mc:Ignorable="missing"'],
    ["<w:p/>", 'mc:Ignorable="mc"'],
    ["<w:p/>", 'mc:ProcessContent="x:wrap"'],
    ["<w:p/>", 'mc:MustUnderstand="x"'],
    ["<w:p/>", 'mc:PreserveElements="x:*"'],
    ["<w:p/>", 'mc:PreserveAttributes="x:*"'],
    ['<w:p><mc:Choice Requires="w"/></w:p>', ""],
    [
      "<w:p><mc:AlternateContent><mc:Fallback/></mc:AlternateContent></w:p>",
      "",
    ],
    ["<w:p><mc:AlternateContent><mc:Choice/></mc:AlternateContent></w:p>", ""],
    [
      '<w:p><mc:AlternateContent><mc:Choice Requires="missing"/></mc:AlternateContent></w:p>',
      "",
    ],
    [
      '<w:p><mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback/><mc:Fallback/></mc:AlternateContent></w:p>',
      "",
    ],
    [
      '<w:p><mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback/><mc:Choice Requires="w"/></mc:AlternateContent></w:p>',
      "",
    ],
  ])(
    "fails explicitly for unsupported or malformed MCE: %s %s",
    (body, attrs) => {
      expect(() => process(body, attrs)).toThrow();
    }
  );

  it("keeps entity values, text whitespace, and foreign attribute namespaces during serialization", () => {
    const out = process(
      '<w:p><w:r x:flag="a&amp;b"><w:t xml:space="preserve"> a &amp; b </w:t></w:r></w:p>'
    );
    const doc = parseXml(out);
    expect(doc.getElementsByTagNameNS(W_NS, "t")[0].textContent).toBe(
      " a & b "
    );
    expect(
      doc
        .getElementsByTagNameNS(W_NS, "r")[0]
        .getAttributeNS("urn:unknown", "flag")
    ).toBe("a&b");
  });
});
