// @vitest-environment jsdom
import type { Mark, Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { exportErrorCode } from "../__testing__/docx";
import { parseXml, R_NS } from "../ooxml/xml";
import { docxSchema } from "../schema";
import type { ExportRefs } from "./exportRefs";
import type { LinkTargets } from "./hyperlink";
import { buildParagraph, NO_IMPORT_SOURCES } from "./importParagraph";
import { NO_IMAGE_REFS } from "./media";
import { serializeParagraph } from "./serializeParagraph";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** Opens a `<w:p>` fragment as a paragraph node */
function open(xml: string, links?: LinkTargets): PMNode {
  const wrapped = parseXml(`<w:wrap ${W_NS} xmlns:r="${R_NS}">${xml}</w:wrap>`);
  const el = wrapped.documentElement.firstElementChild;
  if (!el) throw new Error("no element");
  const node = buildParagraph(
    el,
    0,
    links ? { ...NO_IMPORT_SOURCES, links } : NO_IMPORT_SOURCES
  );
  if (!node) throw new Error("the paragraph could not be modelled");
  return node;
}

const ID_PR = '<w:sdtPr><w:id w:val="7"/></w:sdtPr>';
const PREFIX = `<w:sdt>${ID_PR}`;

function sdtMark(sdtPrefix: string, locked = false, sdtKey = 0): Mark {
  return docxSchema.marks.sdt.create({ sdtPrefix, sdtKey, locked });
}

function runMark(rPr: string | null = null): Mark {
  return docxSchema.marks.run.create({ rPr });
}

function paragraph(...inline: PMNode[]): PMNode {
  return docxSchema.nodes.paragraph.create(null, inline);
}

const wr = (text: string, rPr = "") =>
  `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;

const content = (inner: string) =>
  `${PREFIX}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

describe("a content control around a stretch of text", () => {
  it("puts the runs that share one control inside a single sdt", () => {
    const control = sdtMark(PREFIX, true);
    const node = paragraph(
      docxSchema.text("2026", [control, runMark()]),
      docxSchema.text("AD", [control, runMark("<w:rPr><w:b/></w:rPr>")])
    );
    expect(serializeParagraph(node)).toBe(
      `<w:p>${content(wr("2026") + wr("AD", "<w:rPr><w:b/></w:rPr>"))}</w:p>`
    );
  });

  it("leaves the runs on either side of it outside", () => {
    const node = paragraph(
      docxSchema.text("signed on ", [runMark()]),
      docxSchema.text("2026-08-04", [sdtMark(PREFIX, true), runMark()]),
      docxSchema.text(" final", [runMark()])
    );
    expect(serializeParagraph(node)).toBe(
      `<w:p>${wr("signed on ")}${content(wr("2026-08-04"))}${wr(" final")}</w:p>`
    );
  });

  it("two neighbouring controls stay two controls", () => {
    const other = '<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr>';
    const node = paragraph(
      docxSchema.text("a", [sdtMark(PREFIX), runMark()]),
      docxSchema.text("b", [sdtMark(other), runMark()])
    );
    expect(serializeParagraph(node)).toBe(
      `<w:p>${content(wr("a"))}${other}<w:sdtContent>${wr("b")}` +
        "</w:sdtContent></w:sdt></w:p>"
    );
  });

  it("two controls written exactly alike stay two controls as well", () => {
    const node = paragraph(
      docxSchema.text("a", [sdtMark(PREFIX), runMark()]),
      // Nothing but which control it is tells this one from the one before it
      docxSchema.text("b", [sdtMark(PREFIX, false, 1), runMark()])
    );
    expect(serializeParagraph(node)).toBe(
      `<w:p>${content(wr("a"))}${content(wr("b"))}</w:p>`
    );
  });

  it("a preserved run inside the control goes back inside it", () => {
    const empty = '<w:r><w:rPr><w:rtl w:val="0"/></w:rPr></w:r>';
    const xml = `<w:p>${content(empty + wr("value"))}</w:p>`;
    expect(serializeParagraph(open(xml))).toBe(xml);
  });

  it("refuses to export a control that has lost its opening XML", () => {
    const node = paragraph(
      docxSchema.text("value", [
        docxSchema.marks.sdt.create({ locked: true }),
        runMark(),
      ])
    );
    expect(exportErrorCode(() => serializeParagraph(node))).toBe(
      "lost-original"
    );
  });
});

describe("a paragraph nobody edited", () => {
  it.each([
    `<w:p>${content(wr("2026-08-04"))}</w:p>`,
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${wr("signed on ")}` +
      `${content(wr("2026") + wr("AD", "<w:rPr><w:b/></w:rPr>"))}${wr(" final")}</w:p>`,
    `<w:p>${content("<w:r><w:tab/><w:br/></w:r>")}</w:p>`,
    `<w:p><w:bookmarkStart w:id="0" w:name="here"/>${content(wr("value"))}</w:p>`,
    `<w:p>${content(
      '<w:proofErr w:type="spellStart"/><w:bookmarkStart w:id="1" w:name="here"/>' +
        `${wr("value")}<w:bookmarkEnd w:id="1"/>`
    )}</w:p>`,
    `<w:p>${content(wr("a"))}${content(wr("b"))}</w:p>`,
  ])("goes back out as the XML it came in as: %s", (xml) => {
    expect(serializeParagraph(open(xml))).toBe(xml);
  });
});

describe("tabs", () => {
  it("writes a marked text tab with its source attributes", () => {
    const node = paragraph(
      docxSchema.text("a", [runMark()]),
      docxSchema.text("\t", [
        runMark(),
        docxSchema.marks.tab.create({ tabAttrs: 'w:custom="keep"' }),
      ]),
      docxSchema.text("b", [runMark()])
    );

    expect(serializeParagraph(node)).toBe(
      '<w:p><w:r><w:t xml:space="preserve">a</w:t>' +
        '<w:tab w:custom="keep"/><w:t xml:space="preserve">b</w:t>' +
        "</w:r></w:p>"
    );
  });

  it("defensively writes an unmarked tab character as w:tab", () => {
    expect(serializeParagraph(paragraph(docxSchema.text("a\tb")))).toBe(
      '<w:p><w:r><w:t xml:space="preserve">a</w:t><w:tab/>' +
        '<w:t xml:space="preserve">b</w:t></w:r></w:p>'
    );
  });
});

/** A link mark as the import builds one, with no relationship read for it */
function linkMark(attrs: {
  linkPrefix?: string | null;
  href?: string | null;
  linkKey?: number;
}): Mark {
  return docxSchema.marks.link.create(attrs);
}

/** The relationships an export hands out: one address, on the id given */
function linkRefs(entries: Record<string, string>): ExportRefs {
  return {
    images: NO_IMAGE_REFS,
    links: {
      relIdOf: (href, was) =>
        was !== null && entries[was] === href
          ? was
          : Object.keys(entries).find((id) => entries[id] === href),
    },
  };
}

describe("a hyperlink around a stretch of text", () => {
  const OPENING = '<w:hyperlink r:id="rId9" w:history="1">';
  const LINKS: LinkTargets = new Map([["rId9", "https://example.com"]]);
  const REFS = linkRefs({ rId9: "https://example.com" });

  it("puts the runs that share one link inside a single hyperlink", () => {
    const link = linkMark({ linkPrefix: OPENING, href: "https://example.com" });
    const node = paragraph(
      docxSchema.text("our ", [link, runMark()]),
      docxSchema.text("terms", [link, runMark("<w:rPr><w:b/></w:rPr>")])
    );
    expect(serializeParagraph(node, REFS)).toBe(
      `<w:p>${OPENING}${wr("our ")}` +
        `${wr("terms", "<w:rPr><w:b/></w:rPr>")}</w:hyperlink></w:p>`
    );
  });

  it("two links written alike stay two hyperlinks", () => {
    const first = linkMark({
      linkPrefix: OPENING,
      href: "https://example.com",
    });
    const second = linkMark({
      linkPrefix: OPENING,
      href: "https://example.com",
      linkKey: 1,
    });
    const node = paragraph(
      docxSchema.text("a", [first, runMark()]),
      docxSchema.text("b", [second, runMark()])
    );
    expect(serializeParagraph(node, REFS)).toBe(
      `<w:p>${OPENING}${wr("a")}</w:hyperlink>` +
        `${OPENING}${wr("b")}</w:hyperlink></w:p>`
    );
  });

  it("points the link at the relationship its address now lives on", () => {
    const link = linkMark({ linkPrefix: OPENING, href: "https://example.org" });
    const node = paragraph(docxSchema.text("terms", [link, runMark()]));
    expect(
      serializeParagraph(
        node,
        linkRefs({ rId9: "https://example.com", rId12: "https://example.org" })
      )
    ).toContain('<w:hyperlink r:id="rId12" w:history="1">');
  });

  it("gives a link the editor made an opening tag naming the namespace", () => {
    const link = linkMark({ href: "https://example.com" });
    const node = paragraph(docxSchema.text("terms", [link, runMark()]));
    expect(serializeParagraph(node, REFS)).toBe(
      `<w:p><w:hyperlink xmlns:r="${R_NS}" r:id="rId9">` +
        `${wr("terms")}</w:hyperlink></w:p>`
    );
  });

  it("puts the address on an anchor-only link the same way, so it supersedes the anchor", () => {
    const link = linkMark({
      linkPrefix: '<w:hyperlink w:anchor="chapter3">',
      href: "https://example.com",
    });
    const node = paragraph(docxSchema.text("terms", [link, runMark()]));
    expect(serializeParagraph(node, REFS)).toContain(
      `<w:hyperlink xmlns:r="${R_NS}" r:id="rId9" w:anchor="chapter3">`
    );
  });

  it("refuses to export a link the editor made with no relationship to point at", () => {
    const link = linkMark({ href: "https://example.com" });
    const node = paragraph(docxSchema.text("terms", [link, runMark()]));
    expect(exportErrorCode(() => serializeParagraph(node))).toBe(
      "unsupported-content"
    );
  });

  it("refuses to export a link with neither an address nor an opening tag", () => {
    const node = paragraph(docxSchema.text("terms", [linkMark({}), runMark()]));
    expect(exportErrorCode(() => serializeParagraph(node))).toBe(
      "lost-original"
    );
  });

  it.each([
    `<w:p><w:hyperlink r:id="rId9" w:history="1" w:tooltip="Our terms">${wr("terms")}</w:hyperlink></w:p>`,
    `<w:p><w:hyperlink w:anchor="chapter3">${wr("Chapter Three")}</w:hyperlink></w:p>`,
    `<w:p>${wr("see ")}<w:hyperlink r:id="rId9">${wr("our")}${wr("terms", "<w:rPr><w:b/></w:rPr>")}</w:hyperlink>${wr(".")}</w:p>`,
    `<w:p>${content(`<w:hyperlink r:id="rId9">${wr("terms")}</w:hyperlink>`)}</w:p>`,
  ])("goes back out as the XML it came in as: %s", (xml) => {
    expect(serializeParagraph(open(xml, LINKS), REFS)).toBe(xml);
  });
});
