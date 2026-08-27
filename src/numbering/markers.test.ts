// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import type { NumberingRef } from "../model/format";
import { computeMarkers } from "./markers";
import { type Numbering, parseNumbering } from "./parseNumbering";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

interface LevelSpec {
  format: string;
  text: string;
  start?: number;
  hanging?: number;
}

function levelXml(ilvl: number, spec: LevelSpec): string {
  const ind =
    spec.hanging === undefined
      ? ""
      : `<w:pPr><w:ind w:left="720" w:hanging="${spec.hanging}"/></w:pPr>`;
  return (
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="${spec.start ?? 1}"/>` +
    `<w:numFmt w:val="${spec.format}"/><w:lvlText w:val="${spec.text}"/>` +
    `${ind}</w:lvl>`
  );
}

/** Builds and parses a numbering.xml holding a single list of levels. Its numId is 1 */
function oneList(levels: LevelSpec[], overrides = ""): Numbering {
  return parseNumbering(
    `<w:numbering ${W_NS}>` +
      `<w:abstractNum w:abstractNumId="0">` +
      levels.map((spec, ilvl) => levelXml(ilvl, spec)).join("") +
      "</w:abstractNum>" +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/>${overrides}</w:num>` +
      "</w:numbering>"
  );
}

/** Writes out a sequence of paragraphs using level numbers alone */
function items(...ilvls: number[]): (NumberingRef | null)[] {
  return ilvls.map((ilvl) => ({ numId: 1, ilvl }));
}

function texts(
  paragraphs: readonly (NumberingRef | null)[],
  numbering: Numbering
): (string | null)[] {
  return computeMarkers(paragraphs, numbering).map(
    (marker) => marker?.text ?? null
  );
}

describe("the five number formats", () => {
  const sequence = items(0, 0, 0);

  it.each([
    ["decimal", "%1.", ["1.", "2.", "3."]],
    ["lowerLetter", "%1)", ["a)", "b)", "c)"]],
    ["upperLetter", "%1.", ["A.", "B.", "C."]],
    ["lowerRoman", "%1.", ["i.", "ii.", "iii."]],
    // A bullet has no number to count, so its lvlText is drawn as it is
    ["bullet", "■", ["■", "■", "■"]],
  ])("%s", (format, text, expected) => {
    expect(texts(sequence, oneList([{ format, text }]))).toEqual(expected);
  });

  it("the letter after z carries on the way Word does", () => {
    const many = items(...Array.from({ length: 27 }, () => 0));
    const all = texts(many, oneList([{ format: "lowerLetter", text: "%1" }]));
    expect(all[25]).toBe("z");
    expect(all[26]).toBe("aa");
  });

  it("builds Roman numerals correctly", () => {
    const many = items(...Array.from({ length: 14 }, () => 0));
    const all = texts(many, oneList([{ format: "lowerRoman", text: "%1" }]));
    expect(all[3]).toBe("iv");
    expect(all[8]).toBe("ix");
    expect(all[13]).toBe("xiv");
  });
});

describe("multiple levels", () => {
  const threeLevels = oneList([
    { format: "decimal", text: "%1." },
    { format: "decimal", text: "%1.%2." },
    { format: "decimal", text: "%1.%2.%3." },
  ]);

  it("%n is the current number of the nth level", () => {
    expect(texts(items(0, 1, 1, 2), threeLevels)).toEqual([
      "1.",
      "1.1.",
      "1.2.",
      "1.2.1.",
    ]);
  });

  it("when a shallower level advances the deeper levels count from the start again", () => {
    expect(texts(items(0, 1, 1, 0, 1, 2), threeLevels)).toEqual([
      "1.",
      "1.1.",
      "1.2.",
      "2.",
      "2.1.",
      "2.1.1.",
    ]);
  });

  it("a shallower level that has not appeared yet reads as its start number", () => {
    expect(texts(items(1), threeLevels)).toEqual(["1.1."]);
  });

  it("keeps a parenthesized shape as it is", () => {
    const numbering = oneList([
      { format: "decimal", text: "%1." },
      { format: "decimal", text: "(%2)" },
    ]);
    expect(texts(items(0, 1, 1), numbering)).toEqual(["1.", "(1)", "(2)"]);
  });
});

describe("start number", () => {
  it("uses start as the first number", () => {
    expect(
      texts(
        items(0, 0),
        oneList([{ format: "decimal", text: "%1.", start: 3 }])
      )
    ).toEqual(["3.", "4."]);
  });

  it("startOverride beats start", () => {
    const numbering = oneList(
      [{ format: "decimal", text: "%1.", start: 3 }],
      '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride>'
    );
    expect(texts(items(0, 0), numbering)).toEqual(["7.", "8."]);
  });

  /**
   * A crafted file can ask for a start no real list has, and a letter or roman marker
   * grows with the number it spells, so the length of the marker is bounded as well
   */
  describe("a start number far past what a list would use", () => {
    const bomb = 2_000_000_000;

    it.each(["lowerLetter", "upperLetter", "lowerRoman"])(
      "draws %s as a decimal instead of spelling it out",
      (format) => {
        const numbering = oneList([{ format, text: "%1.", start: bomb }]);
        expect(texts(items(0), numbering)).toEqual([`${bomb}.`]);
      }
    );

    it("spells out the numbers a list really uses", () => {
      const numbering = oneList([
        { format: "lowerLetter", text: "%1.", start: 26 },
      ]);
      expect(texts(items(0, 0), numbering)).toEqual(["z.", "aa."]);
    });
  });
});

/**
 * The shape a level lays down is drawn as it is written, so bounding the number a marker spells
 * bounds nothing on its own: a crafted `w:lvlText` of a megabyte would be drawn in front of every
 * paragraph of the list, and drawn again on every keystroke. The marker itself is cut to the
 * length README records instead.
 */
describe("a level text longer than a marker is ever drawn", () => {
  const MARKER_CHARS = 64;
  const long = "shape ".repeat(20_000);

  it("is cut to length, number and all", () => {
    const numbering = oneList([{ format: "decimal", text: `${long}%1.` }]);
    const marker = computeMarkers(items(0), numbering)[0];

    expect(marker?.text).toHaveLength(MARKER_CHARS);
    expect(marker?.text).toBe(long.slice(0, MARKER_CHARS));
  });

  it("is cut where it is a bullet as well, which is drawn as it stands", () => {
    const numbering = oneList([{ format: "bullet", text: long }]);
    expect(computeMarkers(items(0), numbering)[0]?.text).toHaveLength(
      MARKER_CHARS
    );
  });

  it("is cut between characters, never through one", () => {
    // A code-unit cut lands in the middle of the surrogate pair after the leading letter
    const shape = `x${"\u{1f642}".repeat(20_000)}`;
    const numbering = oneList([{ format: "bullet", text: shape }]);
    const marker = computeMarkers(items(0), numbering)[0];

    expect(marker?.text).toBe(`x${"\u{1f642}".repeat(31)}`);
    expect(marker?.text.length).toBeLessThanOrEqual(MARKER_CHARS);
  });

  it("keeps a sequence that draws as one character whole", () => {
    const family = "\u{1f468}\u200d\u{1f469}\u200d\u{1f467}";
    const numbering = oneList([{ format: "bullet", text: family.repeat(100) }]);
    const marker = computeMarkers(items(0), numbering)[0];

    expect(marker?.text).toBe(family.repeat(8));
    expect(marker?.text.length).toBeLessThanOrEqual(MARKER_CHARS);
  });

  it("draws no marker at all where not even one character fits the cap", () => {
    // A single base letter carrying thousands of combining marks is one character to cut at
    const numbering = oneList([
      { format: "bullet", text: `a${"\u0301".repeat(20_000)}` },
    ]);
    expect(computeMarkers(items(0), numbering)[0]).toBeNull();
  });

  it("leaves the shapes a real list writes as they are", () => {
    const numbering = oneList([{ format: "decimal", text: "Article %1." }]);
    expect(texts(items(0), numbering)).toEqual(["Article 1."]);
  });
});

describe("paragraphs that are not lists", () => {
  const numbering = oneList([{ format: "decimal", text: "%1." }]);

  it("a paragraph outside a list is passed over without counting a number", () => {
    expect(texts([...items(0), null, ...items(0)], numbering)).toEqual([
      "1.",
      null,
      "2.",
    ]);
  });

  it("a level the list does not have gets no number", () => {
    expect(texts([{ numId: 1, ilvl: 5 }], numbering)).toEqual([null]);
  });

  it("counts the numbers separately for each list", () => {
    const two = parseNumbering(
      `<w:numbering ${W_NS}>` +
        '<w:abstractNum w:abstractNumId="0">' +
        levelXml(0, { format: "decimal", text: "%1." }) +
        "</w:abstractNum>" +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
        '<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>' +
        "</w:numbering>"
    );
    const paragraphs = [
      { numId: 1, ilvl: 0 },
      { numId: 2, ilvl: 0 },
      { numId: 1, ilvl: 0 },
    ];
    expect(texts(paragraphs, two)).toEqual(["1.", "1.", "2."]);
  });
});

describe("hanging indent", () => {
  it("returns the indent the level sets along with it", () => {
    const numbering = oneList([
      { format: "decimal", text: "%1.", hanging: 360 },
    ]);
    expect(computeMarkers(items(0), numbering)[0]).toEqual({
      text: "1.",
      indent: { startPt: 36, endPt: null, textIndentPt: -18 },
    });
  });

  it("is empty when the level sets no indent", () => {
    const numbering = oneList([{ format: "decimal", text: "%1." }]);
    expect(computeMarkers(items(0), numbering)[0]).toEqual({
      text: "1.",
      indent: { startPt: null, endPt: null, textIndentPt: null },
    });
  });
});

/**
 * A list started fresh while editing has to show its numbers on screen even though its
 * definition is not in the document
 */
describe("a list number the document does not know", () => {
  const numbering = oneList([{ format: "decimal", text: "%1." }]);

  it("an even number is drawn with the numbered list template", () => {
    const paragraphs = [
      { numId: 20, ilvl: 0 },
      { numId: 20, ilvl: 1 },
      { numId: 20, ilvl: 1 },
      { numId: 20, ilvl: 2 },
      { numId: 20, ilvl: 0 },
    ];
    expect(texts(paragraphs, numbering)).toEqual([
      "1.",
      "a.",
      "b.",
      "i.",
      "2.",
    ]);
  });

  it("an odd number is drawn with the bullet list template", () => {
    const paragraphs = [
      { numId: 21, ilvl: 0 },
      { numId: 21, ilvl: 1 },
      { numId: 21, ilvl: 2 },
    ];
    expect(texts(paragraphs, numbering)).toEqual(["●", "○", "■"]);
  });

  it("every template level has a hanging indent", () => {
    expect(computeMarkers([{ numId: 20, ilvl: 3 }], numbering)[0]).toEqual({
      text: "1.",
      indent: { startPt: 144, endPt: null, textIndentPt: -18 },
    });
  });
});
