// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { parseXml, W_NS } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { fieldSpans } from "./fields";
import { NO_FORMATTING } from "./formatting";
import { NO_IMPORT_SOURCES } from "./importParagraph";
import { buildBlock } from "./story";

/** One paragraph read the way a story block is, which is where a field span is read off */
function paragraph(inner: string) {
  const el = parseXml(`<w:p xmlns:w="${W_NS}">${inner}</w:p>`).documentElement;
  const node = buildBlock(el, "", NO_IMPORT_SOURCES, NO_FORMATTING);
  if (node.type !== docxSchema.nodes.paragraph) {
    throw new Error("the fixture did not read as a paragraph");
  }
  return node;
}

const BEGIN = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
const SEPARATE = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

function instruction(text: string): string {
  return `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
}

describe("field spans", () => {
  it("pairs a complex field's begin, separate and end", () => {
    const node = paragraph(
      "<w:r><w:t>a</w:t></w:r>" +
        BEGIN +
        instruction(" PAGE ") +
        SEPARATE +
        "<w:r><w:t>7</w:t></w:r>" +
        END
    );

    expect(fieldSpans(node, 0)).toEqual([
      { instr: "PAGE", begin: 1, separate: 3, end: 5, result: [4, 5] },
    ]);
  });

  it("counts every piece of an instruction split across runs", () => {
    const node = paragraph(
      BEGIN + instruction(" NUM") + instruction("PAGES ") + SEPARATE + END
    );

    expect(fieldSpans(node, 0)[0]?.instr).toBe("NUMPAGES");
  });

  it("reads a simple field's instruction and cached result off the one element", () => {
    const node = paragraph(
      '<w:fldSimple w:instr=" PAGE  \\* MERGEFORMAT "><w:r><w:t>3</w:t></w:r></w:fldSimple>'
    );

    expect(fieldSpans(node, 0)).toEqual([
      {
        instr: "PAGE  \\* MERGEFORMAT",
        begin: 0,
        separate: null,
        end: 0,
        result: null,
      },
    ]);
  });

  it("leaves a field that caches no result without one", () => {
    const node = paragraph(BEGIN + instruction(" TIME ") + END);

    expect(fieldSpans(node, 0)).toEqual([
      { instr: "TIME", begin: 0, separate: null, end: 2, result: null },
    ]);
  });

  it("passes over a field that never ends", () => {
    const node = paragraph(
      BEGIN + instruction(" PAGE ") + "<w:r><w:t>tail</w:t></w:r>"
    );

    expect(fieldSpans(node, 0)).toEqual([]);
  });

  it("passes over an end with no field to close", () => {
    const node = paragraph("<w:r><w:t>a</w:t></w:r>" + END);

    expect(fieldSpans(node, 0)).toEqual([]);
  });

  it("reads a field nested in another one, in the order they begin", () => {
    const node = paragraph(
      BEGIN +
        instruction(" IF ") +
        SEPARATE +
        BEGIN +
        instruction(" PAGE ") +
        SEPARATE +
        "<w:r><w:t>2</w:t></w:r>" +
        END +
        END
    );

    expect(fieldSpans(node, 0).map((span) => span.instr)).toEqual([
      "IF",
      "PAGE",
    ]);
  });

  it("counts positions from the offset it is given", () => {
    const node = paragraph(BEGIN + instruction(" PAGE ") + SEPARATE + END);

    expect(fieldSpans(node, 10)).toEqual([
      { instr: "PAGE", begin: 10, separate: 12, end: 13, result: [13, 13] },
    ]);
  });
});
