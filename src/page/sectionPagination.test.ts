// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  LETTER_FIXTURE,
  LETTER_LANDSCAPE_SECT_PR,
  LETTER_SECT_PR,
  makeDocx,
  readFixture,
} from "../__testing__/docx";
import { importDocx } from "../docx/importDocx";
import { sectionsOf } from "../docx/sections";
import type { MeasuredBlock } from "./blockKinds";
import { A4_PAGE_PIXELS, pageLayout, sectionPixels } from "./pageLayout";

/**
 * The page count a document written on its own paper comes to, with every block drawn the same
 * height. The heights are the one thing a browser measures, so they are given here and everything
 * else - which section each block belongs to, and the paper that section names - is read out of
 * the document the way the editor reads it.
 */
function pages(doc: PMNode, height: number): number {
  const blocks: MeasuredBlock[] = [];
  doc.forEach((_node, offset) => {
    blocks.push({
      pos: offset,
      gap: 0,
      height,
      breakBefore: false,
      breakAfter: false,
      candidates: [],
      minFirstPiece: height,
      keepWithNext: false,
    });
  });
  return pageLayout({ blocks, sections: sectionPixels(sectionsOf(doc)) }).pages
    .length;
}

const RUN = '<w:r><w:t xml:space="preserve">block</w:t></w:r>';

/** The same Letter paper, declared as a section that starts where the one before it ended */
const LETTER_CONTINUOUS_SECT_PR = LETTER_SECT_PR.replace(
  "<w:sectPr>",
  '<w:sectPr><w:type w:val="continuous"/>'
);

function paragraphs(count: number): string {
  return `<w:p>${RUN}</w:p>`.repeat(count);
}

describe("the pages a document comes to", () => {
  /**
   * Letter is 11in of paper less an inch of margin at either end, which leaves 864px of body
   * where A4 leaves 971px: nine blocks of 100px fill one A4 page and run onto a second Letter one.
   */
  it("follows the paper the document itself names", () => {
    const letter = importDocx(readFixture(LETTER_FIXTURE)).doc;
    const nine = importDocx(makeDocx(paragraphs(9) + LETTER_SECT_PR)).doc;

    expect(A4_PAGE_PIXELS.bodyHeight).toBeCloseTo(971, 0);
    expect(pages(nine, 100)).toBe(2);
    // The same nine blocks on the paper A4 leaves come to one page
    expect(pages(importDocx(makeDocx(paragraphs(9))).doc, 100)).toBe(1);
    // The committed fixture is written on that same Letter paper
    expect(sectionPixels(sectionsOf(letter))[0]?.pixels.bodyHeight).toBeCloseTo(
      864,
      0
    );
  });

  /**
   * Three blocks on Letter upright and seven on Letter sideways. The section break opens the
   * second page, and the sideways paper leaves 624px of body, so the seven blocks run onto a
   * third page where the upright paper's 864px would have held them all on the second.
   */
  it("gives a landscape second section the pages its own paper leaves", () => {
    const twoSections = importDocx(
      makeDocx(
        `${paragraphs(2)}<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>${RUN}</w:p>` +
          paragraphs(7) +
          LETTER_LANDSCAPE_SECT_PR
      )
    ).doc;

    expect(sectionsOf(twoSections)).toHaveLength(2);
    expect(pages(twoSections, 100)).toBe(3);

    // On one paper for the whole document those ten blocks come to two pages
    const oneSection = importDocx(
      makeDocx(paragraphs(10) + LETTER_SECT_PR)
    ).doc;
    expect(pages(oneSection, 100)).toBe(2);
  });

  /**
   * A continuous section starts where the one before it ended (§17.6.22), so eight blocks of
   * 100px still come to the one Letter page their 800px fit on, break or no break.
   */
  it("does not open a page for a section that declares itself continuous", () => {
    const continuous = importDocx(
      makeDocx(
        `${paragraphs(2)}<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>${RUN}</w:p>` +
          paragraphs(5) +
          LETTER_CONTINUOUS_SECT_PR
      )
    ).doc;

    expect(sectionsOf(continuous)).toHaveLength(2);
    expect(pages(continuous, 100)).toBe(1);
    // The same document whose second section starts on a new page comes to two
    const nextPage = importDocx(
      makeDocx(
        `${paragraphs(2)}<w:p><w:pPr>${LETTER_SECT_PR}</w:pPr>${RUN}</w:p>` +
          paragraphs(5) +
          LETTER_SECT_PR
      )
    ).doc;
    expect(pages(nextPage, 100)).toBe(2);
  });
});
