// @vitest-environment node
import { describe, expect, it } from "vitest";
import { parseAttrs, readTag, rootTagAt } from "./tagScan";

describe("readTag", () => {
  it("steps over a > inside a quoted attribute", () => {
    const source = "<w:p w:a=\"x>y\" w:b='c>d'>text</w:p>";
    const tag = readTag(source, 0);
    expect(tag).toEqual({
      kind: "open",
      name: "w:p",
      nameEnd: 4,
      end: source.indexOf(">text") + 1,
    });
  });

  it("tells the four kinds of construct apart", () => {
    expect(readTag("<w:p/>", 0)?.kind).toBe("empty");
    expect(readTag("</w:p>", 0)?.kind).toBe("close");
    expect(readTag("<!-- <p> -->", 0)?.kind).toBe("other");
    expect(readTag('<?xml version="1.0"?>', 0)?.kind).toBe("other");
    expect(readTag("<![CDATA[<p>]]>", 0)?.kind).toBe("other");
    expect(readTag("<!DOCTYPE p>", 0)?.kind).toBe("other");
  });

  /**
   * A `>` may stand inside a comment or a CDATA section, so each is read to its own closing
   * marker rather than to the first `>`, or the text after it would be read as markup.
   */
  it("reads a comment and a CDATA section to their own end, past a > inside them", () => {
    const comment = "<!-- a > b -->";
    expect(readTag(`${comment}<w:p/>`, 0)).toEqual({
      kind: "other",
      name: "",
      nameEnd: comment.length,
      end: comment.length,
    });
    const cdata = "<![CDATA[ a > b ]]>";
    expect(readTag(`${cdata}<w:p/>`, 0)).toEqual({
      kind: "other",
      name: "",
      nameEnd: cdata.length,
      end: cdata.length,
    });
    const declaration = '<?xml version="1.0"?>';
    expect(readTag(`${declaration}<w:p/>`, 0)?.end).toBe(declaration.length);
  });

  it("is null for a tag that never ends and for one with no name", () => {
    expect(readTag('<w:p w:a="1"', 0)).toBeNull();
    expect(readTag("<!-- unterminated", 0)).toBeNull();
    expect(readTag("< >", 0)).toBeNull();
  });
});

describe("parseAttrs", () => {
  it("reads the pairs of an attribute string under the names they were written with", () => {
    expect(parseAttrs("w:val=\"single\" w14:x='2'")).toEqual([
      ["w:val", "single"],
      ["w14:x", "2"],
    ]);
    expect(parseAttrs("")).toEqual([]);
  });

  it("decodes the five entities and numeric references in attribute values", () => {
    expect(
      parseAttrs('w:val="&amp;&lt;&gt;&quot;&apos;" w:x="&#65;&#x42;&#x1F600;"')
    ).toEqual([
      ["w:val", "&<>\"'"],
      ["w:x", "AB\u{1F600}"],
    ]);
  });

  it("is null for a shape a parser would refuse", () => {
    expect(parseAttrs("w:val=single")).toBeNull();
    expect(parseAttrs('w:val="1" w:val="2"')).toBeNull();
    expect(parseAttrs('w:val="1"w:x="2"')).toBeNull();
    expect(parseAttrs('w:val="a&unknown;b"')).toBeNull();
    expect(parseAttrs('w:val="a&#xD800;"')).toBeNull();
    expect(parseAttrs("w:val")).toBeNull();
  });
});

describe("rootTagAt", () => {
  it("finds the root past a prolog and a comment that holds a <", () => {
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      "<!-- <w:people> was here -->" +
      "<w15:people/>";
    expect(rootTagAt(xml)).toBe(xml.indexOf("<w15:people/>"));
  });

  it("is -1 for a source that holds no element", () => {
    expect(rootTagAt("<!-- nothing else -->")).toBe(-1);
    expect(rootTagAt("")).toBe(-1);
    expect(rootTagAt("<!-- unterminated")).toBe(-1);
  });
});
