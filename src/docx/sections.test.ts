// @vitest-environment jsdom

import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  decode,
  LETTER_GEOMETRY as LETTER,
  LETTER_SECT_PR,
  makeDocx,
  readFixture,
} from "../__testing__/docx";
import { childByLocalName, parseXml } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { onlyCommentsChangedBy } from "./commentOnlyChange";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import { A4_PORTRAIT } from "./pageGeometry";
import {
  DEFAULT_SECTION,
  firstSectPrElement,
  parseSectionProperties,
  readSectionProperties,
  sectionAt,
  sectionIn,
  sectionsOf,
  setSectionChild,
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

function opened(body: string) {
  return importDocx(makeDocx(body));
}

function documentXmlOf(bytes: Uint8Array): string {
  return decode(unzipSync(bytes)["word/document.xml"]);
}

/** Where the paragraph holding exactly this text stands, at any depth, and its last position */
function paragraphAt(doc: PMNode, text: string): { pos: number; to: number } {
  const found: { pos: number; to: number }[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === "paragraph" && node.textContent === text) {
      found.push({ pos, to: pos + node.nodeSize - 1 });
    }
    return found.length === 0;
  });
  const first = found[0];
  if (first === undefined) throw new Error(`no paragraph reading "${text}"`);
  return first;
}

/**
 * The sections every position of the document reaches, asserting the sections tile the document
 * with no gap and no overlap and that each position is answered by the section holding it
 */
function sectionsReachedIn(
  doc: PMNode,
  sections: ReturnType<typeof sectionsOf>
) {
  expect(sections[0].from).toBe(0);
  expect(sections[sections.length - 1].to).toBe(doc.content.size);
  sections.slice(1).forEach((section, i) => {
    expect(section.from).toBe(sections[i].to + 1);
  });
  const reached = new Set<number>();
  for (let pos = 0; pos <= doc.content.size; pos += 1) {
    const section = sectionIn(sections, pos);
    const holding = sections.filter(
      (candidate) => candidate.from <= pos && pos <= candidate.to
    );
    expect(holding).toEqual([section]);
    reached.add(section.index);
  }
  return reached;
}

describe("the sections a document is written in", () => {
  it("reads one section per pPr sectPr plus the body sectPr", () => {
    const sections = sectionsOf(opened(TWO_SECTIONS).doc);

    expect(sections.map((section) => section.index)).toEqual([0, 1]);
    expect(sections[0].props.geometry).toEqual(LETTER);
    expect(sections[1].props.geometry).toEqual({
      widthTwips: 16838,
      heightTwips: 11906,
      marginLeftTwips: 1134,
      marginRightTwips: 1134,
      marginTopTwips: 1134,
      marginBottomTwips: 1134,
    });
    expect(sections[1].props.pageNumberStart).toBe(5);
    expect(sections[1].props.type).toBe("nextPage");
  });

  it("anchors the last section to the body and the others to their paragraphs", () => {
    const { doc } = opened(TWO_SECTIONS);
    const sections = sectionsOf(doc);

    const anchor = sections[0].anchor;
    expect(anchor).toEqual({ kind: "paragraph", pos: 0 });
    expect(sections[0].from).toBe(0);
    expect(sections[0].to).toBe(doc.child(0).nodeSize - 1);
    expect(sections[1].anchor).toEqual({ kind: "body" });
    expect(sections[1].from).toBe(doc.child(0).nodeSize);
    expect(sections[1].to).toBe(doc.content.size);
    // The position the anchor names is the paragraph that closes the section
    if (anchor.kind !== "paragraph") throw new Error("no paragraph anchor");
    expect(doc.resolve(anchor.pos + 1).parent.textContent).toBe(
      "first section"
    );
    expect(sectionAt(doc, 1).index).toBe(0);
    expect(sectionAt(doc, doc.content.size - 1).index).toBe(1);
  });

  /** The rule `spec/notes/sections.md` states for a break written inside a control */
  it("reads a section break carried by a paragraph inside a content control", () => {
    const { doc } = opened(
      `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>` +
        "<w:p><w:r><w:t>inside</w:t></w:r></w:p>" +
        `<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>` +
        "<w:r><w:t>ends the section</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>past the break</w:t></w:r></w:p>" +
        "</w:sdtContent></w:sdt>" +
        "<w:p><w:r><w:t>after the control</w:t></w:r></w:p>" +
        A4_LANDSCAPE_SECT_PR
    );
    const control = doc.child(0);
    expect(control.type.name).toBe("sdtBlock");

    const sections = sectionsOf(doc);
    expect(sections).toHaveLength(2);
    expect(sections[0].props.geometry).toEqual(LETTER);
    // The anchor names the paragraph inside the control, where the break is written
    const anchor = sections[0].anchor;
    if (anchor.kind !== "paragraph") throw new Error("no paragraph anchor");
    const ends = paragraphAt(doc, "ends the section");
    expect(anchor.pos).toBe(ends.pos);
    // The section ends at that paragraph, not at the control holding it
    expect(sections[0].to).toBe(ends.to);
    expect(sections[0].to).toBeLessThan(control.nodeSize - 1);
    expect(sections[1].from).toBe(sections[0].to + 1);
    expect(sectionAt(doc, anchor.pos + 1).index).toBe(0);
    // The paragraph standing after the break inside the same control is already the next section
    expect(
      sectionAt(doc, paragraphAt(doc, "past the break").pos + 1).index
    ).toBe(1);
    expect(sectionAt(doc, control.nodeSize + 1).index).toBe(1);
  });

  it("a control holding no break does not part the section it stands in", () => {
    const { doc } = opened(
      "<w:p><w:r><w:t>before</w:t></w:r></w:p>" +
        `<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr><w:sdtContent>` +
        "<w:p><w:r><w:t>inside</w:t></w:r></w:p>" +
        "</w:sdtContent></w:sdt>" +
        "<w:p><w:r><w:t>after</w:t></w:r></w:p>" +
        A4_LANDSCAPE_SECT_PR
    );

    const sections = sectionsOf(doc);
    expect(sections).toHaveLength(1);
    expect(sections[0].anchor).toEqual({ kind: "body" });
  });

  /** Both breaks are read, and every position still reaches exactly one section */
  it("reads both breaks a control holds", () => {
    const { doc } = opened(
      "<w:p><w:r><w:t>before</w:t></w:r></w:p>" +
        `<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent>` +
        `<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>` +
        "<w:r><w:t>ends the first</w:t></w:r></w:p>" +
        `<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>` +
        "<w:r><w:t>ends the second</w:t></w:r></w:p>" +
        "</w:sdtContent></w:sdt>" +
        "<w:p><w:r><w:t>after</w:t></w:r></w:p>" +
        A4_LANDSCAPE_SECT_PR
    );
    const sections = sectionsOf(doc);

    expect(sections).toHaveLength(3);
    expect(sections[0].props.geometry).toEqual(LETTER);
    expect(sections[1].props.geometry).toEqual(LETTER);
    expect(sections.map((section) => section.anchor)).toEqual([
      { kind: "paragraph", pos: paragraphAt(doc, "ends the first").pos },
      { kind: "paragraph", pos: paragraphAt(doc, "ends the second").pos },
      { kind: "body" },
    ]);

    // The second section is ended by the paragraph straight after the first's, so it covers that
    // paragraph and nothing else
    const second = paragraphAt(doc, "ends the second");
    expect(sections[1].from).toBe(second.pos);
    expect(sections[1].to).toBe(second.to);

    // Every position reaches exactly one section, and every section is reached
    for (const section of sections) {
      expect(section.from).toBeLessThanOrEqual(section.to);
    }
    expect(sectionsReachedIn(doc, sections)).toEqual(new Set([0, 1, 2]));
  });

  /** `eachStoryParagraph` reaches a control inside a control, so a break written there is read */
  it("reads a break written inside a control nested in a control", () => {
    const { doc } = opened(
      `<w:sdt><w:sdtPr><w:id w:val="10"/></w:sdtPr><w:sdtContent>` +
        "<w:p><w:r><w:t>outer</w:t></w:r></w:p>" +
        `<w:sdt><w:sdtPr><w:id w:val="11"/></w:sdtPr><w:sdtContent>` +
        `<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>` +
        "<w:r><w:t>ends the section</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>past the break</w:t></w:r></w:p>" +
        "</w:sdtContent></w:sdt>" +
        "</w:sdtContent></w:sdt>" +
        "<w:p><w:r><w:t>after</w:t></w:r></w:p>" +
        A4_LANDSCAPE_SECT_PR
    );
    const sections = sectionsOf(doc);

    expect(sections).toHaveLength(2);
    const ends = paragraphAt(doc, "ends the section");
    expect(doc.resolve(ends.pos).parent.type.name).toBe("sdtBlock");
    expect(sections[0].to).toBe(ends.to);
    expect(
      sectionAt(doc, paragraphAt(doc, "past the break").pos + 1).index
    ).toBe(1);
    expect(sectionsReachedIn(doc, sections)).toEqual(new Set([0, 1]));
  });

  /**
   * A cell's paragraphs are not the sequence the body ends its sections in, so a `w:sectPr`
   * written in one stays preserved markup rather than becoming a section.
   */
  it("reads no section off a paragraph inside a table cell", () => {
    const { doc } = opened(
      "<w:tbl><w:tr><w:tc>" +
        `<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>` +
        "<w:r><w:t>in a cell</w:t></w:r></w:p>" +
        "</w:tc></w:tr></w:tbl>" +
        A4_LANDSCAPE_SECT_PR
    );

    expect(sectionsOf(doc)).toHaveLength(1);
  });

  it("reads no section off a paragraph inside a table cell inside a content control", () => {
    const { doc } = opened(
      `<w:sdt><w:sdtPr><w:id w:val="12"/></w:sdtPr><w:sdtContent>` +
        "<w:tbl><w:tr><w:tc>" +
        `<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>` +
        "<w:r><w:t>in a cell</w:t></w:r></w:p>" +
        "</w:tc></w:tr></w:tbl>" +
        "</w:sdtContent></w:sdt>" +
        A4_LANDSCAPE_SECT_PR
    );
    expect(doc.child(0).type.name).toBe("sdtBlock");
    expect(doc.child(0).child(0).type.name).toBe("table");

    const sections = sectionsOf(doc);
    expect(sections.map((section) => section.anchor)).toEqual([
      { kind: "body" },
    ]);
  });

  /** The one committed package holding a section break written inside a control */
  it("reads the break the content controls fixture writes inside a control", () => {
    const { doc } = importDocx(readFixture("content-controls.docx"));
    const sections = sectionsOf(doc);

    expect(sections).toHaveLength(2);
    const anchor = sections[0].anchor;
    if (anchor.kind !== "paragraph") throw new Error("no paragraph anchor");
    expect(doc.resolve(anchor.pos).parent.type.name).toBe("sdtBlock");
    // The two sections are told apart by the margins each one is written to
    expect(sections[0].props.geometry.marginTopTwips).not.toBe(
      sections[1].props.geometry.marginTopTwips
    );
    // Everything after the control belongs to the section the body closes
    expect(sectionAt(doc, sections[0].to + 1).index).toBe(1);
  });

  it("the final sectPr goes back out byte for byte from the doc attribute", () => {
    const { doc, session } = opened(TWO_SECTIONS);

    expect(doc.attrs.sectPr).toBe(A4_LANDSCAPE_SECT_PR);
    // Nothing of the body's own section is left among the blocks or in the preserved tail
    expect(doc.childCount).toBe(2);
    expect(session.documentSuffix).not.toContain("w:sectPr");
    expect(documentXmlOf(exportDocx(doc, session))).toContain(
      `</w:p>${A4_LANDSCAPE_SECT_PR}</w:body>`
    );
  });

  it("a document with no sectPr has one A4 section", () => {
    const { doc } = opened("<w:p><w:r><w:t>only text</w:t></w:r></w:p>");
    const sections = sectionsOf(doc);

    expect(doc.attrs.sectPr).toBeNull();
    expect(sections).toHaveLength(1);
    expect(sections[0].props).toEqual(DEFAULT_SECTION);
    expect(sections[0].props.geometry).toEqual(A4_PORTRAIT);
    expect(sections[0].anchor).toEqual({ kind: "body" });
  });

  it("the session's first-section geometry equals sectionsOf(doc)[0]", () => {
    const { doc, session } = opened(TWO_SECTIONS);

    // The import reads the paper off the body's DOM before there is a document node to read it
    // off, so the two readings have to be the same reading
    expect(session.geometry).toEqual(sectionsOf(doc)[0].props.geometry);
    expect(session.geometry).toEqual(LETTER);
  });

  it("reads a page-number start no counter could reach as none at all", () => {
    // `1` with 309 zeroes after it is `Infinity` once read, and a page counter starting there
    // would draw a page number nothing could follow
    const { doc } = opened(
      `<w:p/><w:sectPr><w:pgNumType w:start="1e309"/>` +
        '<w:type w:val="sideways"/></w:sectPr>'
    );
    const section = sectionsOf(doc)[0];

    expect(section.props.pageNumberStart).toBeNull();
    // §17.6.22 names five kinds, and a section claiming another names none
    expect(section.props.type).toBeNull();
  });

  it("setSectionChild inserts pgMar after pgSz by CT_SectPr order", () => {
    const narrow =
      '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>';
    const written = setSectionChild(
      '<w:sectPr><w:pgSz w:w="16838" w:h="11906"/><w:cols w:num="2"/></w:sectPr>',
      "pgMar",
      narrow
    );

    expect(written).toBe(
      '<w:sectPr><w:pgSz w:w="16838" w:h="11906"/>' +
        `${narrow}<w:cols w:num="2"/></w:sectPr>`
    );
  });

  it("setSectionChild replaces a repeated headerReference by type", () => {
    const section =
      "<w:sectPr>" +
      '<w:headerReference w:type="default" r:id="rId4"/>' +
      '<w:headerReference w:type="first" r:id="rId5"/>' +
      '<w:footerReference w:type="default" r:id="rId6"/>' +
      "</w:sectPr>";

    expect(
      setSectionChild(
        section,
        "headerReference",
        '<w:headerReference w:type="first" r:id="rId9"/>'
      )
    ).toBe(
      "<w:sectPr>" +
        '<w:headerReference w:type="default" r:id="rId4"/>' +
        '<w:headerReference w:type="first" r:id="rId9"/>' +
        '<w:footerReference w:type="default" r:id="rId6"/>' +
        "</w:sectPr>"
    );
    // A variant the section did not name is added rather than written over another one
    expect(
      setSectionChild(
        section,
        "headerReference",
        '<w:headerReference w:type="even" r:id="rId9"/>'
      )
    ).toContain(
      '<w:headerReference w:type="first" r:id="rId5"/>' +
        '<w:headerReference w:type="even" r:id="rId9"/>'
    );
  });

  it("splitting the last paragraph of a section leaves the break on the later half", () => {
    const pPr = `<w:pPr><w:jc w:val="center"/>${LETTER_SECT_PR}</w:pPr>`;

    expect(withoutSectionBreak(pPr)).toBe(
      '<w:pPr><w:jc w:val="center"/></w:pPr>'
    );
    // Properties holding nothing else collapse rather than going out empty
    expect(withoutSectionBreak(`<w:pPr>${LETTER_SECT_PR}</w:pPr>`)).toBeNull();
    // A section the split moved is still one section, read off the halves the split leaves
    const { doc } = opened(TWO_SECTIONS);
    const paragraph = doc.child(0);
    const split = docxSchema.nodes.doc.create(doc.attrs, [
      paragraph.type.create(
        { ...paragraph.attrs, pPr: withoutSectionBreak(paragraph.attrs.pPr) },
        paragraph.content
      ),
      paragraph,
      doc.child(1),
    ]);

    const halves = sectionsOf(split);
    expect(halves.map((section) => section.from)).toEqual([
      0,
      split.child(0).nodeSize + split.child(1).nodeSize,
    ]);
    expect(halves[0].to).toBe(
      split.child(0).nodeSize + split.child(1).nodeSize - 1
    );
  });
});

describe("surgical edits to interleaved section references", () => {
  const headerDefault = '<w:headerReference w:type="default" r:id="rId4"/>';
  const footerDefault = '<w:footerReference w:type="default" r:id="rId5"/>';
  const headerFirst = '<w:headerReference w:type="first" r:id="rId6"/>';
  const section =
    `<w:sectPr>\n${headerDefault}\n<!-- footer -->${footerDefault}` +
    `\n<!-- first page -->${headerFirst}\n<w:pgSz w:w="12240" w:h="15840"/>` +
    "\n<!-- section tail --></w:sectPr>";

  it.each([
    ["headerReference", headerFirst],
    ["footerReference", footerDefault],
  ])(
    "replaces a %s in place without moving siblings or gaps",
    (name, original) => {
      const replacement = original.replace(/rId\d+/, "rId9");
      expect(setSectionChild(section, name, replacement)).toBe(
        section.replace(original, replacement)
      );
    }
  );

  it.each([
    ["headerReference", headerFirst],
    ["footerReference", footerDefault],
  ])(
    "adds a new %s variant after its kind without moving existing references",
    (name, previous) => {
      const added = `<w:${name} w:type="even" r:id="rId9"/>`;
      expect(setSectionChild(section, name, added)).toBe(
        section.replace(previous, previous + added)
      );
    }
  );
});

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
    // Tolerate a missing required variant as the default one.
    expect(props.footerRefs.default).toBe("rId6");
    expect(props.titlePg).toBe(true);
  });

  it("keeps the text a section was written as when it is read out of a fragment", () => {
    const props = parseSectionProperties(A4_LANDSCAPE_SECT_PR);

    expect(props?.xml).toBe(A4_LANDSCAPE_SECT_PR);
    expect(props?.geometry.widthTwips).toBe(16838);
    expect(parseSectionProperties("<w:pPr/>")).toBeNull();
    expect(parseSectionProperties("not xml at all")).toBeNull();
  });

  it("puts a section break back into the spot CT_PPr gives it", () => {
    const pPr = `<w:pPr><w:jc w:val="center"/>${LETTER_SECT_PR}</w:pPr>`;

    expect(
      withSectionBreak('<w:pPr><w:jc w:val="center"/></w:pPr>', LETTER_SECT_PR)
    ).toBe(pPr);
    expect(withSectionBreak(null, LETTER_SECT_PR)).toBe(
      `<w:pPr>${LETTER_SECT_PR}</w:pPr>`
    );
  });
});

it("ignores foreign section names before the Word section", () => {
  const foreign =
    '<x:sectPr xmlns:x="urn:foreign"><x:pgSz x:w="1" x:h="2"/></x:sectPr>';
  const { session } = importDocx(
    makeDocx(
      `<w:p><w:r><w:drawing>${foreign}</w:drawing></w:r></w:p>${LETTER_SECT_PR}`
    )
  );
  expect(session.geometry).toEqual(LETTER);
});

it("edits a body section after its preserved comment gap", () => {
  const gap = "\n<!-- before <w:sectPr> -->\n";
  const child = '<w:pgMar w:top="100"/>';
  expect(setSectionChild(gap + LETTER_SECT_PR, "pgMar", child)).toBe(
    gap + LETTER_SECT_PR.replace(/<w:pgMar[^>]*\/>/, child)
  );
});

it("models a section-only body without inventing XML until its paragraph is edited", () => {
  const input = makeDocx(LETTER_SECT_PR);
  const { doc, session } = importDocx(input);
  expect(doc.attrs.sectPr).toBe(LETTER_SECT_PR);
  expect(sectionsOf(doc)[0].props.geometry).toEqual(session.geometry);
  expect(doc.firstChild?.type.name).toBe("paragraph");
  doc.check();
  const original = decode(unzipSync(input)["word/document.xml"]);
  expect(decode(unzipSync(exportDocx(doc, session))["word/document.xml"])).toBe(
    original
  );
  expect(onlyCommentsChangedBy(input, exportDocx(doc, session), "me")).toEqual({
    ok: true,
  });
  const state = EditorState.create({ doc });
  const edited = state.apply(state.tr.insertText("added", 1)).doc;
  const output = decode(
    unzipSync(exportDocx(edited, session))["word/document.xml"]
  );
  expect(importDocx(exportDocx(edited, session)).doc.textContent).toBe("added");
  expect(output).toContain(LETTER_SECT_PR + "</w:body>");
});
