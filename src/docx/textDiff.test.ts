import { describe, expect, it } from "vitest";
import { diffText, type TextEdit } from "./textDiff";

function spelled(
  edits: readonly TextEdit[],
  side: "original" | "revised"
): string {
  const dropped = side === "original" ? "added" : "removed";
  return edits
    .filter((edit) => edit.kind !== dropped)
    .map((edit) => edit.text)
    .join("");
}

const graphemes = new Intl.Segmenter(undefined, { granularity: "grapheme" });

function charactersOf(text: string): string[] {
  return Array.from(graphemes.segment(text), ({ segment }) => segment);
}

function editLength(edits: readonly TextEdit[]): number {
  return edits
    .filter((edit) => edit.kind !== "kept")
    .reduce((length, edit) => length + charactersOf(edit.text).length, 0);
}

/** The shortest edit's length by the quadratic table, to hold the bisection against */
function shortestEditLength(original: string, revised: string): number {
  const a = charactersOf(original);
  const b = charactersOf(revised);
  const common = Array.from({ length: a.length + 1 }, () =>
    new Array<number>(b.length + 1).fill(0)
  );
  for (let i = a.length - 1; i >= 0; i -= 1) {
    for (let j = b.length - 1; j >= 0; j -= 1) {
      common[i][j] =
        a[i] === b[j]
          ? common[i + 1][j + 1] + 1
          : Math.max(common[i + 1][j], common[i][j + 1]);
    }
  }
  return a.length + b.length - 2 * common[0][0];
}

describe("diffText", () => {
  it("answers two empty texts with no edit", () => {
    expect(diffText("", "")).toEqual([]);
  });

  it("inserts everything into an empty text and deletes everything out of one", () => {
    expect(diffText("", "abc")).toEqual([{ kind: "added", text: "abc" }]);
    expect(diffText("abc", "")).toEqual([{ kind: "removed", text: "abc" }]);
  });

  it("answers identical texts with one equal stretch", () => {
    expect(diffText("same text", "same text")).toEqual([
      { kind: "kept", text: "same text" },
    ]);
  });

  it("keeps a shared prefix and suffix around the edit", () => {
    expect(diffText("The cat sat", "The dog sat")).toEqual([
      { kind: "kept", text: "The " },
      { kind: "removed", text: "cat" },
      { kind: "added", text: "dog" },
      { kind: "kept", text: " sat" },
    ]);
  });

  it("finds an edit as short as any, where several are", () => {
    // Myers' own example: the shortest edit script between these is five characters long
    const edits = diffText("abcabba", "cbabac");
    expect(spelled(edits, "original")).toBe("abcabba");
    expect(spelled(edits, "revised")).toBe("cbabac");
    expect(editLength(edits)).toBe(5);
  });

  it("is as short as the quadratic table says over many small texts", () => {
    const alphabet = ["a", "b", "c", "가", "😀", "e\u0301"];
    let seed = 7;
    // xorshift32, which keeps its low bits where a floating-point generator loses them
    const next = () => {
      seed ^= seed << 13;
      seed ^= seed >>> 17;
      seed ^= seed << 5;
      return seed >>> 0;
    };
    const lengths = new Set<number>();
    const text = () => {
      const length = next() % 12;
      lengths.add(length);
      return Array.from(
        { length },
        () => alphabet[next() % alphabet.length]
      ).join("");
    };
    for (let round = 0; round < 500; round += 1) {
      const original = text();
      const revised = text();
      const edits = diffText(original, revised);
      expect(spelled(edits, "original")).toBe(original);
      expect(spelled(edits, "revised")).toBe(revised);
      expect(editLength(edits)).toBe(shortestEditLength(original, revised));
    }
    expect(lengths.size).toBe(12);
  });

  it("edits Hangul syllable by syllable", () => {
    expect(diffText("계약 기간은 1년이다", "계약 기간은 2년이다")).toEqual([
      { kind: "kept", text: "계약 기간은 " },
      { kind: "removed", text: "1" },
      { kind: "added", text: "2" },
      { kind: "kept", text: "년이다" },
    ]);
    expect(diffText("갑과 을", "갑과 병")).toEqual([
      { kind: "kept", text: "갑과 " },
      { kind: "removed", text: "을" },
      { kind: "added", text: "병" },
    ]);
  });

  it("never splits a surrogate pair", () => {
    // Both emoji share their high surrogate, so a diff over UTF-16 units would keep half of each
    expect(diffText("a😀b", "a😁b")).toEqual([
      { kind: "kept", text: "a" },
      { kind: "removed", text: "😀" },
      { kind: "added", text: "😁" },
      { kind: "kept", text: "b" },
    ]);
  });

  it("never splits a flag, a skin-tone modifier, or a combining accent from its letter", () => {
    expect(diffText("🇰🇷", "🇰🇵")).toEqual([
      { kind: "removed", text: "🇰🇷" },
      { kind: "added", text: "🇰🇵" },
    ]);
    expect(diffText("ok 👍🏻", "ok 👍🏽")).toEqual([
      { kind: "kept", text: "ok " },
      { kind: "removed", text: "👍🏻" },
      { kind: "added", text: "👍🏽" },
    ]);
    expect(diffText("cafe\u0301", "cafe\u0300")).toEqual([
      { kind: "kept", text: "caf" },
      { kind: "removed", text: "e\u0301" },
      { kind: "added", text: "e\u0300" },
    ]);
  });

  it("spells both texts back and merges neighbouring stretches of one kind", () => {
    const pairs: readonly (readonly [string, string])[] = [
      ["kitten sitting", "sitting kitten"],
      ["abcdefghij", "axcxexgxix"],
      ["완전히 다른 문장", "a totally different line"],
      ["🙂🙃 x", "x 🙃🙂"],
    ];
    for (const [original, revised] of pairs) {
      const edits = diffText(original, revised);
      expect(spelled(edits, "original")).toBe(original);
      expect(spelled(edits, "revised")).toBe(revised);
      expect(edits.every((edit) => edit.text !== "")).toBe(true);
      expect(
        edits.every((edit, at) => at === 0 || edits[at - 1].kind !== edit.kind)
      ).toBe(true);
    }
  });

  it("finds the edit in a long text without comparing it throughout", () => {
    const original = "word ".repeat(20_000);
    const revised = `${"word ".repeat(10_000)}new ${"word ".repeat(10_000)}`;
    expect(diffText(original, revised)).toEqual([
      { kind: "kept", text: "word ".repeat(10_000) },
      { kind: "added", text: "new " },
      { kind: "kept", text: "word ".repeat(10_000) },
    ]);
  });

  it("rewrites two long texts sharing nothing into one removal and one addition", () => {
    const original = "a".repeat(5_000);
    const revised = "b".repeat(5_000);
    expect(diffText(original, revised)).toEqual([
      { kind: "removed", text: original },
      { kind: "added", text: revised },
    ]);
  });

  it("falls back to one removal and one addition between the shared ends of two long unrelated texts", () => {
    let seed = 11;
    const letter = (alphabet: string) => {
      seed ^= seed << 13;
      seed ^= seed >>> 17;
      seed ^= seed << 5;
      return alphabet[(seed >>> 0) % alphabet.length];
    };
    const words = (alphabet: string) =>
      Array.from({ length: 30_000 }, () => letter(alphabet)).join("");
    const original = `Start ${words("abcdefghijklm ")} end.`;
    const revised = `Start ${words("abcdefghijklm ")} end.`;
    const started = performance.now();
    const edits = diffText(original, revised);
    expect(performance.now() - started).toBeLessThan(1_000);
    expect(edits).toEqual([
      { kind: "kept", text: expect.stringMatching(/^Start /) },
      { kind: "removed", text: expect.any(String) },
      { kind: "added", text: expect.any(String) },
      { kind: "kept", text: expect.stringMatching(/ end\.$/) },
    ]);
    expect(spelled(edits, "original")).toBe(original);
    expect(spelled(edits, "revised")).toBe(revised);
  });
});
