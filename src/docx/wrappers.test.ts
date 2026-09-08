// @vitest-environment jsdom

import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { decode, documentXmlOf, makeLinkedDocx } from "../__testing__/docx";
import { parseXml, R_NS } from "../ooxml/xml";
import { wrapperMarks } from "../schema/wrappers";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import type { LinkTargets } from "./hyperlink";
import { importDocx } from "./importDocx";
import { buildParagraph, NO_IMPORT_SOURCES } from "./importParagraph";
import { serializeParagraph } from "./serializeParagraph";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const TERMS = "https://example.com/terms";
const LINKS: LinkTargets = new Map([["rId9", TERMS]]);

/** The relationships an export hands out: the one address, on the id it came in on */
const REFS: ExportRefs = {
  ...NO_EXPORT_REFS,
  links: { relIdOf: (href) => (href === TERMS ? "rId9" : undefined) },
};

function open(xml: string): PMNode {
  const wrapped = parseXml(`<w:wrap ${W_NS} xmlns:r="${R_NS}">${xml}</w:wrap>`);
  const el = wrapped.documentElement.firstElementChild;
  if (!el) throw new Error("no element");
  return buildParagraph(el, null, { ...NO_IMPORT_SOURCES, links: LINKS });
}

const run = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const control = (inner: string, id = 7) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr>` +
  `<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

const link = (inner: string) =>
  `<w:hyperlink r:id="rId9">${inner}</w:hyperlink>`;

/** The one inline node a paragraph holding nothing else opened as */
function only(node: PMNode): PMNode {
  expect(node.type.name).toBe("paragraph");
  expect(node.childCount).toBe(1);
  return node.child(0);
}

/** What every wrapper the node stands inside is, outermost first */
function nesting(node: PMNode): string[] {
  return wrapperMarks(node).map((mark) => mark.type.name);
}

describe("a wrapper holding a wrapper", () => {
  it("a hyperlink holding a control round-trips with the link outside", () => {
    const xml = `<w:p>${link(control(run("terms")))}</w:p>`;
    const node = open(xml);

    expect(nesting(only(node))).toEqual(["link", "sdt"]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("a control holding a link round-trips with the control outside", () => {
    const xml = `<w:p>${control(link(run("terms")))}</w:p>`;
    const node = open(xml);

    expect(nesting(only(node))).toEqual(["sdt", "link"]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("two nested controls keep their order and both ids", () => {
    const xml = `<w:p>${control(control(run("terms"), 8))}</w:p>`;
    const node = open(xml);
    const marks = wrapperMarks(only(node));

    expect(marks.map((mark) => mark.attrs.sdtPrefix)).toEqual([
      '<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr>',
      '<w:sdt><w:sdtPr><w:id w:val="8"/></w:sdtPr>',
    ]);
    // Two controls written alike would still be two, so each carries its own number
    expect(marks.map((mark) => mark.attrs.key)).toEqual([0, 1]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("three wrappers deep still rebuild in file order", () => {
    const xml = `<w:p>${control(link(control(run("terms"), 8)))}</w:p>`;
    const node = open(xml);

    expect(nesting(only(node))).toEqual(["sdt", "link", "sdt"]);
    expect(wrapperMarks(only(node)).map((mark) => mark.attrs.depth)).toEqual([
      0, 1, 2,
    ]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("keeps the text beside a nested wrapper outside it", () => {
    const xml = `<w:p>${run("a")}${link(`${run("b")}${control(run("c"))}${run("d")}`)}${run("e")}</w:p>`;
    const node = open(xml);

    expect(node.textContent).toBe("abcde");
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });
});

/**
 * `w:customXml` wraps inline content exactly as a control does, and no kind is registered for it,
 * so the paragraph around it stays editable and the wrapper is kept whole where it stood
 * (`docx/importPolicy`). It is not the whole paragraph that is stood down.
 */
describe("a wrapper kind nobody registered", () => {
  it("is kept whole inside a paragraph that stays editable", () => {
    const xml = `<w:p>${run("a")}<w:customXml w:element="x">${run("b")}</w:customXml></w:p>`;
    const node = open(xml);

    expect(node.type.name).toBe("paragraph");
    expect(node.children.map((child) => child.type.name)).toEqual([
      "text",
      "rawInline",
    ]);
    expect(node.child(1).attrs.element).toBe("customXml");
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });

  it("keeps the wrappers around it, so they still close where they did", () => {
    const xml = `<w:p>${link(`<w:customXml w:element="x">${run("b")}</w:customXml>`)}</w:p>`;
    const node = open(xml);

    expect(nesting(only(node))).toEqual(["link"]);
    expect(serializeParagraph(node, REFS)).toBe(xml);
  });
});

describe("a document holding a hyperlink that holds a control", () => {
  const BODY = `<w:p>${link(control(run("terms")))}</w:p>`;

  it("opens as an editable paragraph and exports byte identical", () => {
    const bytes = makeLinkedDocx(BODY, { rId9: TERMS });
    const { doc, session } = importDocx(bytes);

    expect(doc.child(0).type.name).toBe("paragraph");
    expect(documentXmlOf(doc, session)).toBe(
      decode(unzipSync(bytes)["word/document.xml"])
    );
  });
});
