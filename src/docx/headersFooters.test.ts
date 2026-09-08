// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  makeHeadersFootersDocx,
} from "../__testing__/docx";
import { docxSchema } from "../schema";
import { storyNodeOf } from "../schema/stories";
import { exportDocx } from "./exportDocx";
import {
  displayPageNumber,
  type HeadersFooters,
  headerFooterOn,
  headerFooterText,
  variantsFor,
} from "./headersFooters";
import { importDocx } from "./importDocx";
import { sectionsOf } from "./sections";
import type { SessionStore } from "./session";

const encoder = new TextEncoder();
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** The stories one section of the opened document shows, which is what the preview draws */
function sectionStories(
  doc: PMNode,
  session: SessionStore,
  index = 0
): HeadersFooters {
  const section = sectionsOf(doc)[index];
  if (!section) throw new Error(`the document has no section ${index}`);
  return variantsFor(section, session.headerFooterStories, (key) =>
    storyNodeOf(doc, key)
  );
}

function openedStories(bytes: Uint8Array, index = 0): HeadersFooters {
  const { doc, session } = importDocx(bytes);
  return sectionStories(doc, session, index);
}

/** What one visual page's header or footer reads as, and null where the section declares none */
function shown(
  stories: HeadersFooters,
  kind: "headers" | "footers",
  page: number,
  totalPages: number
): string | null {
  const content = headerFooterOn(stories[kind], stories, page);
  if (content === null) return null;
  return headerFooterText(
    content.story,
    displayPageNumber(stories, page),
    totalPages
  );
}

/**
 * The same package with a second section, closed by its own paragraph, naming header3 where the
 * first section names header1. §17.6.17: the section a paragraph closes is written in its `w:pPr`.
 */
function twoSections(): Uint8Array {
  const parts = unzipSync(makeHeadersFootersDocx());
  parts["word/document.xml"] = encoder.encode(
    decode(parts["word/document.xml"]).replace(
      "<w:p><w:r><w:t>Body</w:t></w:r></w:p>",
      "<w:p><w:pPr><w:sectPr" +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<w:headerReference w:type="default" r:id="rId12"/>' +
        '<w:footerReference w:type="default" r:id="rId15"/>' +
        "</w:sectPr></w:pPr><w:r><w:t>First section</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>Second section</w:t></w:r></w:p>"
    )
  );
  return zipSync(parts);
}

describe("header and footer stories", () => {
  it("selects first, default, and even stories using the displayed page number", () => {
    const stories = openedStories(makeHeadersFootersDocx());

    expect(displayPageNumber(stories, 1)).toBe(4);
    expect(shown(stories, "headers", 1, 3)).toBe("First header");
    expect(shown(stories, "headers", 2, 3)).toBe("Default 5 of 3");
    expect(shown(stories, "headers", 3, 3)).toBe("Even header");
    expect(shown(stories, "footers", 1, 3)).toBe("First footer");
  });

  it("draws the same character for w:cr and w:noBreakHyphen as the body does", () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    parts["word/header2.xml"] = encoder.encode(
      `<w:hdr xmlns:w="${W_NS}"><w:p><w:r>` +
        '<w:t xml:space="preserve">re</w:t><w:noBreakHyphen/>' +
        '<w:t xml:space="preserve">read</w:t><w:cr/>' +
        '<w:softHyphen/><w:t xml:space="preserve">again</w:t>' +
        "</w:r></w:p></w:hdr>"
    );

    expect(shown(openedStories(zipSync(parts)), "headers", 1, 3)).toBe(
      "re‑read\nagain"
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

    const stories = openedStories(zipSync(parts));
    expect(displayPageNumber(stories, 1)).toBe(1);
    expect(displayPageNumber(stories, 2)).toBe(2);
  });

  it("reads the first paragraph's direct alignment", () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    parts["word/header1.xml"] = encoder.encode(
      `<w:hdr xmlns:w="${W_NS}"><w:p><w:pPr><w:jc w:val="right"/></w:pPr>` +
        "<w:r><w:t>Right header</w:t></w:r></w:p></w:hdr>"
    );

    const stories = openedStories(zipSync(parts));
    expect(headerFooterOn(stories.headers, stories, 2)?.align).toBe("right");
  });

  it("leaves a selected but undeclared first or even story blank", () => {
    const stories = openedStories(makeHeadersFootersDocx());
    const missing: HeadersFooters = {
      ...stories,
      headers: { ...stories.headers, first: null, even: null },
    };

    expect(shown(missing, "headers", 1, 3)).toBeNull();
    expect(shown(missing, "headers", 3, 3)).toBeNull();
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

    const stories = openedStories(zipSync(parts));
    expect(stories.headers.default).not.toBeNull();
    expect(stories.footers.first).not.toBeNull();
    expect(stories.evenAndOdd).toBe(true);
  });

  it('reads w:val="off" on evenAndOddHeaders as off', () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    parts["word/settings.xml"] = encoder.encode(
      `<w:settings xmlns:w="${W_NS}"><w:evenAndOddHeaders w:val="off"/></w:settings>`
    );

    expect(openedStories(zipSync(parts)).evenAndOdd).toBe(false);
    // Every other spelling of off says the same, and the element on its own still says on
    for (const [written, expected] of [
      ['<w:evenAndOddHeaders w:val="0"/>', false],
      ['<w:evenAndOddHeaders w:val="false"/>', false],
      ['<w:evenAndOddHeaders w:val="on"/>', true],
      ["<w:evenAndOddHeaders/>", true],
    ] as const) {
      parts["word/settings.xml"] = encoder.encode(
        `<w:settings xmlns:w="${W_NS}">${written}</w:settings>`
      );
      expect(openedStories(zipSync(parts)).evenAndOdd).toBe(expected);
    }
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

    expect(shown(openedStories(zipSync(parts)), "headers", 2, 2)).toBe(
      "Visible"
    );
  });

  it("keeps the cached result of a field it does not work out again", () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    parts["word/header2.xml"] = encoder.encode(
      `<w:hdr xmlns:w="${W_NS}"><w:p>` +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> TITLE </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
        "<w:r><w:t>Quarterly report</w:t></w:r>" +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
        "</w:p></w:hdr>"
    );

    expect(shown(openedStories(zipSync(parts)), "headers", 1, 3)).toBe(
      "Quarterly report"
    );
  });

  it("keeps reading past a field that never ends", () => {
    const parts = unzipSync(makeHeadersFootersDocx());
    parts["word/header2.xml"] = encoder.encode(
      `<w:hdr xmlns:w="${W_NS}"><w:p>` +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
        '<w:r><w:t xml:space="preserve">Draft</w:t></w:r>' +
        "</w:p></w:hdr>"
    );

    // Nothing pairs with the begin character, so no page number is put there; the characters
    // themselves are field plumbing rather than words, and the text beside them still reads
    expect(shown(openedStories(zipSync(parts)), "headers", 1, 3)).toBe("Draft");
  });

  it("a second section selects its own header parts", () => {
    const bytes = twoSections();
    const first = openedStories(bytes, 0);
    const second = openedStories(bytes, 1);

    // The first section names header3 and footer3 as its own default, where the body's section
    // names header1 and footer1 and starts its page numbering at four
    expect(shown(first, "headers", 1, 2)).toBe("Even header");
    expect(shown(first, "footers", 1, 2)).toBe("Even footer");
    expect(shown(second, "headers", 2, 2)).toBe("Default 5 of 2");
    expect(shown(second, "footers", 2, 2)).toBe("Default footer");
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
