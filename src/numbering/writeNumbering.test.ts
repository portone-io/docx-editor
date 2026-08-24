// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { exportErrorCode } from "../__testing__/docx";
import { parseNumbering } from "./parseNumbering";
import { addListDefinitions } from "./writeNumbering";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const ABSTRACT_2 =
  '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0">' +
  '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>';
const NUM_1 = '<w:num w:numId="1"><w:abstractNumId w:val="2"/></w:num>';

const ORIGINAL = `<w:numbering ${W_NS}>${ABSTRACT_2}${NUM_1}</w:numbering>`;

describe("splicing in new list definitions", () => {
  it("leaves the original text as it is when there is nothing to add", () => {
    expect(addListDefinitions(ORIGINAL, [])).toBe(ORIGINAL);
  });

  it("does not change a single character of the original", () => {
    const written = addListDefinitions(ORIGINAL, [2]);
    // The definition is spliced in before the first number and the number before the
    // closing tag, so the original survives as three chunks
    const firstNumAt = ORIGINAL.indexOf("<w:num ");
    const closeAt = ORIGINAL.lastIndexOf("</w:numbering>");
    const head = ORIGINAL.slice(0, firstNumAt);
    const middle = ORIGINAL.slice(firstNumAt, closeAt);
    const tail = ORIGINAL.slice(closeAt);

    expect(written.startsWith(head)).toBe(true);
    expect(written.endsWith(tail)).toBe(true);
    const between = written.slice(head.length, written.length - tail.length);
    const middleAt = between.indexOf(middle);
    expect(middleAt).toBeGreaterThanOrEqual(0);

    // Put back together, the three chunks are the original again, character for character
    expect(
      head + between.slice(middleAt, middleAt + middle.length) + tail
    ).toBe(ORIGINAL);
    expect(written.length).toBeGreaterThan(ORIGINAL.length);
  });

  it("the definition goes before the first number and the number goes at the very end", () => {
    const written = addListDefinitions(ORIGINAL, [2]);
    const definitionAt = written.indexOf('w:abstractNumId="3"');
    expect(definitionAt).toBeGreaterThan(0);
    expect(definitionAt).toBeLessThan(written.indexOf(NUM_1));
    expect(written.indexOf('<w:num w:numId="2"')).toBeGreaterThan(
      written.indexOf(NUM_1)
    );
  });

  it("puts the number before an element that is required to come last", () => {
    const withCleanup = `<w:numbering ${W_NS}>${ABSTRACT_2}${NUM_1}<w:numIdMacAtCleanup w:val="5"/></w:numbering>`;
    const written = addListDefinitions(withCleanup, [2]);
    expect(written.indexOf('<w:num w:numId="2"')).toBeLessThan(
      written.indexOf("<w:numIdMacAtCleanup")
    );
  });

  it("adds it even to a document that had no definitions at all", () => {
    const empty = `<w:numbering ${W_NS}></w:numbering>`;
    const written = addListDefinitions(empty, [2]);
    expect(parseNumbering(written).lists.get(2)?.levels.size).toBe(9);
  });

  it("stops when there is no closing tag", () => {
    expect(
      exportErrorCode(() => addListDefinitions(`<w:numbering ${W_NS}/>`, [2]))
    ).toBe("malformed-xml");
  });
});

describe("reading the spliced-in definitions back", () => {
  const written = addListDefinitions(ORIGINAL, [3, 4]);
  const numbering = parseNumbering(written);

  it("an even number reads as 1. / a. / i.", () => {
    const levels = numbering.lists.get(4)?.levels;
    expect(levels?.get(0)).toEqual({
      format: "decimal",
      text: "%1.",
      start: 1,
      indent: {
        startTwips: 720,
        endTwips: null,
        hangingTwips: 360,
        firstLineTwips: null,
      },
    });
    expect(levels?.get(1)?.format).toBe("lowerLetter");
    expect(levels?.get(2)?.format).toBe("lowerRoman");
    expect(levels?.get(8)?.indent?.startTwips).toBe(6480);
  });

  it("an odd number reads as bullets", () => {
    const levels = numbering.lists.get(3)?.levels;
    expect(levels?.get(0)?.format).toBe("bullet");
    expect([0, 1, 2].map((ilvl) => levels?.get(ilvl)?.text)).toEqual([
      "●",
      "○",
      "■",
    ]);
  });

  it("the new definition ids start after the ids already in use", () => {
    expect(written).toContain('<w:abstractNum w:abstractNumId="3">');
    expect(written).toContain('<w:abstractNum w:abstractNumId="4">');
    expect(written).toContain(
      '<w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>'
    );
    expect(written).toContain(
      '<w:num w:numId="4"><w:abstractNumId w:val="4"/></w:num>'
    );
  });
});
