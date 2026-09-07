// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { importErrorCode } from "../__testing__/docx";
import { DocxImportError } from "./errors";
import {
  escapeXml,
  namespaceDecls,
  parseXml,
  W_NS,
  withXmlParser,
  type XmlParser,
} from "./xml";

describe("escapeXml", () => {
  it("turns the characters XML gives meaning to into entity references", () => {
    expect(escapeXml('<a href="b">&')).toBe(
      "&lt;a href=&quot;b&quot;&gt;&amp;"
    );
  });

  it("leaves single quotes alone, since attribute values are assumed to be wrapped in double quotes", () => {
    expect(escapeXml("it's")).toBe("it's");
  });

  // The text stays Korean here so that the stripping is proven over multibyte UTF-8
  it("strips the control characters XML 1.0 cannot carry", () => {
    expect(escapeXml("가\u0000나\u0008다\u000B라\u000C마\u001F바")).toBe(
      "가나다라마바"
    );
  });

  it("keeps the tab and the newline XML allows", () => {
    expect(escapeXml("a\tb\r\nc")).toBe("a\tb\r\nc");
  });

  it("yields XML that parses again even when control characters come mixed in", () => {
    const xml = `<w:t xmlns:w="${W_NS}">${escapeXml("cont\u0007ract")}</w:t>`;
    expect(parseXml(xml).documentElement.textContent).toBe("contract");
  });
});

/**
 * The parser is a dependency a caller may name, so that a runtime holding no `DOMParser` is
 * turned down by this package with a code of its own rather than by a `ReferenceError` thrown
 * out of the middle of a read.
 */
describe("the parser a read goes through", () => {
  const XML = `<w:t xmlns:w="${W_NS}">fee</w:t>`;

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  /** A parser that answers through the real one and keeps what it was asked to read */
  function recording(): { parser: XmlParser; sources: string[] } {
    const real = new DOMParser();
    const sources: string[] = [];
    return {
      parser: {
        parseFromString: (source, type) => {
          sources.push(source);
          return real.parseFromString(source, type);
        },
      },
      sources,
    };
  }

  it("reads through the parser handed to the scope rather than the global one", () => {
    const { parser, sources } = recording();

    const read = withXmlParser(parser, () => parseXml(XML));

    expect(read.documentElement.textContent).toBe("fee");
    expect(sources).toEqual([XML]);
  });

  /**
   * A `DOMParser` answers markup it cannot read with a `parsererror` document, but a parser is
   * free to throw instead, and an exception of the parser's own reaching the caller is the very
   * thing a caller cannot tell apart from a runtime giving out
   */
  it("reads a throw of the parser's own as markup that did not parse", () => {
    const real = new DOMParser();
    const throwing: XmlParser = {
      parseFromString: (source, type) => {
        const doc = real.parseFromString(source, type);
        if (doc.getElementsByTagName("parsererror").length > 0) {
          throw new SyntaxError("unexpected end of input");
        }
        return doc;
      },
    };

    expect(
      importErrorCode(() => withXmlParser(throwing, () => parseXml("<w:t>")))
    ).toBe("malformed-xml");
    expect(
      withXmlParser(throwing, () => parseXml(XML)).documentElement
    ).not.toBeNull();
  });

  it("refuses with no-xml-parser when no parser is in scope and none is global", () => {
    vi.stubGlobal("DOMParser", undefined);

    expect(importErrorCode(() => parseXml(XML))).toBe("no-xml-parser");
  });

  it("keeps the outer parser through a nested entry naming none", () => {
    const { parser, sources } = recording();
    vi.stubGlobal("DOMParser", undefined);

    withXmlParser(parser, () => withXmlParser(undefined, () => parseXml(XML)));

    expect(sources).toEqual([XML]);
  });

  it("restores the previous parser after work throws", () => {
    const { parser } = recording();
    vi.stubGlobal("DOMParser", undefined);

    expect(() =>
      withXmlParser(parser, () => {
        throw new Error("work gave up");
      })
    ).toThrow("work gave up");

    expect(importErrorCode(() => parseXml(XML))).toBe("no-xml-parser");
  });
});

/**
 * A part is markup from outside, and the parser expands the entities a DTD declares, so a
 * part carrying one would put text on the screen that the part itself does not hold
 */
describe("a part that declares a DTD", () => {
  it("is refused rather than parsed", () => {
    const withEntity =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<!DOCTYPE w:t [<!ENTITY fee "one hundred">]>' +
      `<w:t xmlns:w="${W_NS}">&fee;</w:t>`;
    expect(() => parseXml(withEntity)).toThrow(DocxImportError);
  });

  it("is refused with no entity to expand either", () => {
    const bare = `<!DOCTYPE w:t SYSTEM "w.dtd"><w:t xmlns:w="${W_NS}">fee</w:t>`;
    expect(() => parseXml(bare)).toThrow(DocxImportError);
  });

  it("is looked for in the prolog alone, so the same text in the content parses", () => {
    const text = '<!DOCTYPE w:t [<!ENTITY fee "one hundred">]>';
    const xml = `<w:t xmlns:w="${W_NS}"><![CDATA[${text}]]></w:t>`;
    expect(parseXml(xml).documentElement.textContent).toBe(text);
  });

  /**
   * A comment and a processing instruction may both hold a `<` of their own, so neither the
   * one nor the other is where the prolog ends
   */
  it("is refused with a comment holding a < standing ahead of it", () => {
    const hidden =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      "<!-- < -->" +
      '<!DOCTYPE w:t [<!ENTITY fee "one hundred">]>' +
      `<w:t xmlns:w="${W_NS}">&fee;</w:t>`;
    expect(importErrorCode(() => parseXml(hidden))).toBe("malformed-xml");
  });

  it("is refused with a processing instruction holding a < standing ahead of it", () => {
    const hidden =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      "<?display <b ?>" +
      '<!DOCTYPE w:t [<!ENTITY fee "one hundred">]>' +
      `<w:t xmlns:w="${W_NS}">&fee;</w:t>`;
    expect(importErrorCode(() => parseXml(hidden))).toBe("malformed-xml");
  });
});

describe("a part that declares no DTD", () => {
  it("parses with a comment and a processing instruction ahead of its root", () => {
    const xml =
      '<?xml version="1.0" encoding="UTF-8"?>' +
      "<!-- written by a generator that signs its work with a < -->" +
      "<?display <b ?>" +
      `<w:t xmlns:w="${W_NS}">fee</w:t>`;
    expect(parseXml(xml).documentElement.textContent).toBe("fee");
  });
});

describe("namespaceDecls", () => {
  it("declares every prefix a fragment uses, w to its namespace and the rest to placeholders", () => {
    const decls = namespaceDecls(
      '<w:pPr><w14:paraId w14:val="1"/><m:oMath/></w:pPr>'
    );

    expect(decls).toContain(`xmlns:w="${W_NS}"`);
    expect(decls).toContain('xmlns:w14="urn:docx-editor:w14"');
    expect(decls).toContain('xmlns:m="urn:docx-editor:m"');
  });

  it("declares w even for a fragment that never names it", () => {
    expect(namespaceDecls("<m:oMathPara/>")).toContain(`xmlns:w="${W_NS}"`);
  });

  it("declares the prefix of an attribute that opens the string", () => {
    const attrs = 'w14:paraId="1A2B3C4D" w:rsidR="00A1B2C3"';

    expect(namespaceDecls(attrs)).toContain('xmlns:w14="urn:docx-editor:w14"');
    expect(() =>
      parseXml(`<x ${namespaceDecls(attrs)} ${attrs}/>`)
    ).not.toThrow();
  });

  it("leaves out the two names a document may not redeclare", () => {
    const decls = namespaceDecls('<w:t xml:space="preserve">a</w:t>');

    expect(decls).not.toContain("xmlns:xml=");
    expect(decls).not.toContain("xmlns:xmlns=");
  });
});
