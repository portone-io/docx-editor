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
import { toParagraphFormat } from "../model/format";
import { isBlockControl } from "../schema/controlAttrs";
import type { MeasuredBlock } from "./blockKinds";
import {
  A4_PAGE_PIXELS,
  type PageLayout,
  pageLayout,
  sectionPixels,
} from "./pageLayout";

interface Held {
  node: PMNode;
  pos: number;
}

/** The paragraphs a block stands for, through every content control around them */
function heldParagraphs(node: PMNode, pos: number): Held[] {
  if (!isBlockControl(node)) return [{ node, pos }];
  const held: Held[] = [];
  node.forEach((child, offset) => {
    held.push(...heldParagraphs(child, pos + 1 + offset));
  });
  return held;
}

function keepsNext(held: Held | undefined): boolean {
  return toParagraphFormat(held?.node.attrs.format)?.keepNext === true;
}

/**
 * The layout a document written on its own paper comes to, with every paragraph drawn the same
 * height. The heights are the one thing a browser measures, so they are given here and everything
 * else - which section each piece belongs to, and the paper that section names - is read out of
 * the document the way the editor reads it. A content control offers a place to part it between
 * every two paragraphs it holds, kept closed under one kept with the next, as its measurer does.
 */
function laidOut(doc: PMNode, height: number): PageLayout {
  const blocks: MeasuredBlock[] = [];
  doc.forEach((node, offset) => {
    const held = heldParagraphs(node, offset);
    blocks.push({
      pos: offset,
      gap: 0,
      height: height * held.length,
      breakBefore: false,
      breakAfter: false,
      candidates: held.slice(1).map((paragraph, index) => ({
        at: paragraph.pos,
        offset: height * (index + 1),
        forced: false,
        ...(keepsNext(held[index]) ? { kept: true } : {}),
        repeatHeight: 0,
      })),
      minFirstPiece: height,
      keepWithNext: keepsNext(held.at(-1)),
    });
  });
  return pageLayout({ blocks, sections: sectionPixels(sectionsOf(doc)) });
}

function pages(doc: PMNode, height: number): number {
  return laidOut(doc, height).pages.length;
}

/** The paragraphs the first content control of the body holds */
function heldByControl(doc: PMNode): Held[] {
  let found: Held[] | null = null;
  doc.forEach((node, offset) => {
    if (found === null && isBlockControl(node)) {
      found = heldParagraphs(node, offset);
    }
  });
  if (found === null) throw new Error("the document holds no content control");
  return found;
}

/** The body height of the page a layout opens with this index, down to where it is parted */
function bodyOf(layout: PageLayout, index: number): number {
  const start = layout.pages[index]?.bodyStart ?? 0;
  return (layout.splits[index]?.y ?? layout.bodyHeight) - start;
}

const RUN = '<w:r><w:t xml:space="preserve">block</w:t></w:r>';

/** The same Letter paper, declared as a section that starts where the one before it ended */
const LETTER_CONTINUOUS_SECT_PR = LETTER_SECT_PR.replace(
  "<w:sectPr>",
  '<w:sectPr><w:type w:val="continuous"/>'
);

/** Letter on its side, declared as a section that starts where the one before it ended */
const LETTER_LANDSCAPE_CONTINUOUS_SECT_PR = LETTER_LANDSCAPE_SECT_PR.replace(
  "<w:sectPr>",
  '<w:sectPr><w:type w:val="continuous"/>'
);

function control(body: string): string {
  return `<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>${body}</w:sdtContent></w:sdt>`;
}

function ending(sectPr: string, pPr = ""): string {
  return `<w:p><w:pPr>${pPr}${sectPr}</w:pPr>${RUN}</w:p>`;
}

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
   * A paragraph inside a block-level content control is one of the body's own (§17.5.2.34), so the
   * break it carries ends a section. Where that paragraph is the control's last, the control ends
   * with its section and the next one opens after it.
   */
  it("opens the next section after a control whose last paragraph ends one", () => {
    const held = importDocx(
      makeDocx(
        paragraphs(2) +
          control(paragraphs(1) + ending(LETTER_SECT_PR)) +
          paragraphs(7) +
          LETTER_LANDSCAPE_SECT_PR
      )
    ).doc;

    expect(held.child(2).type.name).toBe("sdtBlock");
    expect(sectionsOf(held)).toHaveLength(2);
    // Three blocks on Letter upright and seven on Letter sideways, as they are with the break
    // written on a loose paragraph
    expect(pages(held, 100)).toBe(3);
    expect(laidOut(held, 100).cuts).toEqual([]);
  });

  /**
   * Word parts the control where a paragraph in the middle of it ends a section: the four
   * paragraphs up to the break fill Letter upright, and the seven after it run onto a second
   * Letter sideways page as they do with the break on a loose paragraph. Laid on one paper, the
   * eleven would come to two pages.
   */
  it("parts a control where a paragraph in the middle of it ends a section", () => {
    const held = importDocx(
      makeDocx(
        paragraphs(2) +
          control(paragraphs(1) + ending(LETTER_SECT_PR) + paragraphs(7)) +
          LETTER_LANDSCAPE_SECT_PR
      )
    ).doc;
    const loose = importDocx(
      makeDocx(
        paragraphs(3) +
          ending(LETTER_SECT_PR) +
          paragraphs(7) +
          LETTER_LANDSCAPE_SECT_PR
      )
    ).doc;
    const layout = laidOut(held, 100);
    const inside = heldByControl(held);

    expect(pages(held, 100)).toBe(pages(loose, 100));
    expect(layout.pages.map((page) => page.section)).toEqual([0, 1, 1]);
    // The control is parted under the paragraph carrying the break, and the page after it is
    // drawn on the sideways paper's shorter body
    expect(layout.cuts[0]?.at).toBe(inside[2]?.pos);
    expect(layout.splits[0]?.forced).toBe(true);
    expect(bodyOf(layout, 0)).toBeCloseTo(864, 0);
    expect(bodyOf(layout, 1)).toBeCloseTo(624, 0);
  });

  it("parts a control where a paragraph in a control inside it ends a section", () => {
    const held = importDocx(
      makeDocx(
        paragraphs(2) +
          control(
            paragraphs(1) +
              control(paragraphs(1) + ending(LETTER_SECT_PR)) +
              paragraphs(7)
          ) +
          LETTER_LANDSCAPE_SECT_PR
      )
    ).doc;
    const loose = importDocx(
      makeDocx(
        paragraphs(4) +
          ending(LETTER_SECT_PR) +
          paragraphs(7) +
          LETTER_LANDSCAPE_SECT_PR
      )
    ).doc;
    const layout = laidOut(held, 100);

    expect(pages(held, 100)).toBe(3);
    expect(pages(held, 100)).toBe(pages(loose, 100));
    expect(layout.cuts[0]?.at).toBe(heldByControl(held)[3]?.pos);
    expect(layout.pages.map((page) => page.section)).toEqual([0, 1, 1]);
  });

  /**
   * A continuous section starting inside a control carries on down the page on the same paper,
   * where seven paragraphs of 100px fit one Letter page, but on sideways paper it opens a page as
   * Word does, parting the control under the paragraph carrying the break.
   */
  it("opens a page inside a control only for a continuous section on other paper", () => {
    const document = (sectPr: string) =>
      importDocx(
        makeDocx(
          paragraphs(2) +
            control(paragraphs(1) + ending(LETTER_SECT_PR) + paragraphs(3)) +
            sectPr
        )
      ).doc;

    const same = laidOut(document(LETTER_CONTINUOUS_SECT_PR), 100);
    expect(same.splits).toEqual([]);
    expect(same.pages.map((page) => page.section)).toEqual([0]);

    const sideways = document(LETTER_LANDSCAPE_CONTINUOUS_SECT_PR);
    const other = laidOut(sideways, 100);
    expect(other.cuts.map((cut) => cut.at)).toEqual([
      heldByControl(sideways)[2]?.pos,
    ]);
    expect(other.splits.map((split) => split.forced)).toEqual([true]);
    expect(other.pages.map((page) => page.section)).toEqual([0, 1]);
  });

  /**
   * Three sections out of one control: Letter upright up to the first break, Letter sideways up to
   * the second, and Letter upright again for the body's own. The seven paragraphs of the middle
   * section take two sideways pages, each numbered within that section.
   */
  it("draws each of three sections two breaks in one control make", () => {
    const held = importDocx(
      makeDocx(
        paragraphs(1) +
          control(
            paragraphs(1) +
              ending(LETTER_SECT_PR) +
              paragraphs(6) +
              ending(LETTER_LANDSCAPE_SECT_PR) +
              paragraphs(2)
          ) +
          LETTER_SECT_PR
      )
    ).doc;
    const layout = laidOut(held, 100);

    expect(sectionsOf(held)).toHaveLength(3);
    expect(
      layout.pages.map((page) => [page.section, page.pageInSection])
    ).toEqual([
      [0, 1],
      [1, 1],
      [1, 2],
      [2, 1],
    ]);
    expect(bodyOf(layout, 1)).toBeCloseTo(624, 0);
    expect(layout.splits.map((split) => split.forced)).toEqual([
      true,
      false,
      true,
    ]);
  });

  /**
   * A paragraph kept with the next one still ends its section there (§17.6.22 opens the page, and
   * no keep can close it), where two paragraphs of 100px under it would otherwise have been kept
   * with it on Letter upright's one page.
   */
  it("parts a control at a section-ending paragraph kept with the next one", () => {
    const held = importDocx(
      makeDocx(
        paragraphs(2) +
          control(
            paragraphs(1) +
              ending(LETTER_SECT_PR, "<w:keepNext/>") +
              paragraphs(2)
          ) +
          LETTER_LANDSCAPE_SECT_PR
      )
    ).doc;
    const layout = laidOut(held, 100);

    expect(layout.cuts.map((cut) => cut.at)).toEqual([
      heldByControl(held)[2]?.pos,
    ]);
    expect(layout.pages.map((page) => page.section)).toEqual([0, 1]);
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
