// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { W_NS } from "../ooxml/names";
import { normalizePackagePrefixes } from "./packagePrefixes";

const encode = (text: string) => new TextEncoder().encode(text);
const decode = (bytes: Uint8Array) => new TextDecoder("utf-8").decode(bytes);

const FOREIGN_MAIN = `<ns0:document xmlns:ns0="${W_NS}"><ns0:body><ns0:p/></ns0:body></ns0:document>`;

const PACKAGE_RELS =
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Target="word/document.xml" Type="urn:x"/></Relationships>';

const CONTENT_TYPES =
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="xml" ContentType="application/xml"/></Types>';

describe("normalizePackagePrefixes", () => {
  /**
   * The media part carries the very text a part would be rewritten from, so what keeps it whole
   * has to be its name rather than anything about its bytes.
   */
  it("rewrites the markup parts and never decodes the media", () => {
    const media = encode(FOREIGN_MAIN);
    const parts = new Map([
      ["word/document.xml", encode(FOREIGN_MAIN)],
      ["word/header1.xml", encode(FOREIGN_MAIN)],
      ["word/media/image1.png", media],
    ]);
    normalizePackagePrefixes(parts);

    expect(decode(parts.get("word/document.xml") ?? new Uint8Array())).toBe(
      `<w:document xmlns:w="${W_NS}"><w:body><w:p/></w:body></w:document>`
    );
    expect(decode(parts.get("word/header1.xml") ?? new Uint8Array())).toContain(
      "<w:p/>"
    );
    expect(parts.get("word/media/image1.png")).toBe(media);
  });

  /**
   * The package relationships and the content types bind namespaces of their own, which this
   * editor neither reads by prefix nor writes, so nothing about them needs a path to be excluded.
   */
  it("hands back the relationships and the content types as the bytes they arrived as", () => {
    const rels = encode(PACKAGE_RELS);
    const types = encode(CONTENT_TYPES);
    const parts = new Map([
      ["_rels/.rels", rels],
      ["[Content_Types].xml", types],
      ["word/document.xml", encode(FOREIGN_MAIN)],
    ]);
    normalizePackagePrefixes(parts);

    expect(parts.get("_rels/.rels")).toBe(rels);
    expect(parts.get("[Content_Types].xml")).toBe(types);
  });

  /**
   * A part this editor never reads may hold bytes no UTF-8 decoder reads. Decoding those turns
   * each of them into U+FFFD, so writing the text back would put that character in the file.
   */
  it("leaves a markup part whose bytes are not UTF-8 as they arrived", () => {
    const latin1 = new Uint8Array([
      ...encode(`<ns0:root xmlns:ns0="${W_NS}"><ns0:v>`),
      0xe9,
      ...encode("</ns0:v></ns0:root>"),
    ]);
    const parts = new Map([["customXml/item2.xml", latin1]]);
    normalizePackagePrefixes(parts);
    expect(parts.get("customXml/item2.xml")).toBe(latin1);
  });

  /**
   * A part the content types declare as XML holds markup whatever its name is, and the export
   * reads it back as XML, so the open spells it like the rest of the package.
   */
  it("rewrites a part the content types declare as XML under a name of its own", () => {
    const types =
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/customXml/item1.data" ContentType="application/xml"/>' +
      "</Types>";
    const parts = new Map([
      ["[Content_Types].xml", encode(types)],
      ["customXml/item1.data", encode(FOREIGN_MAIN)],
    ]);
    normalizePackagePrefixes(parts);

    expect(
      decode(parts.get("customXml/item1.data") ?? new Uint8Array())
    ).toContain("<w:p/>");
  });

  it("keeps the byte order mark a part arrived with", () => {
    const parts = new Map([
      ["word/document.xml", encode(`\u{FEFF}${FOREIGN_MAIN}`)],
    ]);
    normalizePackagePrefixes(parts);
    const rewritten = parts.get("word/document.xml") ?? new Uint8Array();
    expect([...rewritten.slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(decode(rewritten)).toContain("<w:p/>");
  });
});
