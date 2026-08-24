// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { importErrorCode } from "../__testing__/docx";
import { DocxImportError } from "./errors";
import { escapeXml, parseXml, W_NS } from "./xml";

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
