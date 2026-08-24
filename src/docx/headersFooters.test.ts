// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  makeHeadersFootersDocx,
} from "../__testing__/docx";
import { docxSchema } from "../schema";
import { exportDocx } from "./exportDocx";
import {
  displayPageNumber,
  headerFooterAlign,
  headerFooterText,
} from "./headersFooters";
import { importDocx } from "./importDocx";

const encoder = new TextEncoder();
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

describe("header and footer stories", () => {
  it("selects first, default, and even stories using the displayed page number", () => {
    const { session } = importDocx(makeHeadersFootersDocx());
    const stories = session.headersFooters;

    expect(displayPageNumber(stories, 1)).toBe(4);
    expect(headerFooterText(stories.headers, stories, 1, 3)).toBe(
      "First header"
    );
    expect(headerFooterText(stories.headers, stories, 2, 3)).toBe(
      "Default 5 of 3"
    );
    expect(headerFooterText(stories.headers, stories, 3, 3)).toBe(
      "Even header"
    );
    expect(headerFooterText(stories.footers, stories, 1, 3)).toBe(
      "First footer"
    );
  });

  it("starts at page one when pgNumType does not declare a start", () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    parts["word/document.xml"] = encoder.encode(
      decode(parts["word/document.xml"]).replace(
        '<w:pgNumType w:start="4"/>',
        "<w:pgNumType/>"
      )
    );

    const stories = importDocx(zipSync(parts)).session.headersFooters;
    expect(displayPageNumber(stories, 1)).toBe(1);
    expect(displayPageNumber(stories, 2)).toBe(2);
  });

  it("reads the first paragraph's direct alignment", () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    parts["word/header1.xml"] = encoder.encode(
      `<w:hdr xmlns:w="${W_NS}"><w:p><w:pPr><w:jc w:val="right"/></w:pPr>` +
        "<w:r><w:t>Right header</w:t></w:r></w:p></w:hdr>"
    );

    const stories = importDocx(zipSync(parts)).session.headersFooters;
    expect(headerFooterAlign(stories.headers, stories, 2)).toBe("right");
  });

  it("leaves a selected but undeclared first or even story blank", () => {
    const { session } = importDocx(makeHeadersFootersDocx());
    const stories = session.headersFooters;
    const missing = { ...stories.headers, first: null, even: null };

    expect(headerFooterText(missing, stories, 1, 3)).toBeNull();
    expect(headerFooterText(missing, stories, 3, 3)).toBeNull();
  });

  it("reads relationship targets relative to the main part", () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    for (const name of ["header1.xml", "footer2.xml", "settings.xml"]) {
      const original = parts[`word/${name}`];
      if (!original) throw new Error(`missing word/${name}`);
      parts[name] = original;
      delete parts[`word/${name}`];
    }
    const relsPath = "word/_rels/document.xml.rels";
    parts[relsPath] = encoder.encode(
      decode(parts[relsPath])
        .replace('Target="header1.xml"', 'Target="../header1.xml"')
        .replace('Target="footer2.xml"', 'Target="../footer2.xml"')
        .replace('Target="settings.xml"', 'Target="../settings.xml"')
    );

    const stories = importDocx(zipSync(parts)).session.headersFooters;
    expect(stories.headers.default).not.toBeNull();
    expect(stories.footers.first).not.toBeNull();
    expect(stories.evenAndOdd).toBe(true);
  });

  it("shows only top-level paragraph text from a story", () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    parts["word/header1.xml"] = encoder.encode(
      `<w:hdr xmlns:w="${W_NS}">` +
        "<w:p><w:r><w:t>Visible</w:t></w:r>" +
        "<w:r><w:drawing><w:txbxContent><w:p><w:r><w:t>Text box</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r>" +
        "</w:p>" +
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
        "</w:hdr>"
    );

    const stories = importDocx(zipSync(parts)).session.headersFooters;
    expect(headerFooterText(stories.headers, stories, 2, 2)).toBe("Visible");
  });

  it("keeps every related story part byte-identical after a body edit", () => {
    const bytes = makeHeadersFootersDocx();
    const original = unzipSync(bytes);
    const opened = importDocx(bytes);
    const paragraph = opened.doc.firstChild;
    if (!paragraph) throw new Error("the test document has no paragraph");
    const edited = opened.doc.copy(
      opened.doc.content.replaceChild(
        0,
        paragraph.type.create(paragraph.attrs, [docxSchema.text("Changed")])
      )
    );
    const exported = unzipSync(exportDocx(edited, opened.session));

    for (const path of [
      "word/header1.xml",
      "word/header2.xml",
      "word/header3.xml",
      "word/footer1.xml",
      "word/footer2.xml",
      "word/footer3.xml",
      "word/settings.xml",
    ]) {
      expect(bytesEqual(exported[path], original[path]), path).toBe(true);
    }
  });
});
