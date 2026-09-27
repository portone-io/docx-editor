import { describe, expect, it } from "vitest";
import {
  type Alignable,
  alignSequences,
  commonSubsequence,
} from "./sequenceDiff";

interface Item extends Alignable {
  kind: string;
}

const item = (identity: string, text = identity, kind = "p"): Item => ({
  identity,
  text,
  kind,
});

const align = (original: readonly Item[], revised: readonly Item[]) =>
  alignSequences(original, revised, (entry) => entry.kind);

describe("commonSubsequence", () => {
  it("matches the elements of a longest common subsequence in order", () => {
    const original = Array.from("abcabba");
    const revised = Array.from("cbabac");
    const matches = commonSubsequence(
      original.length,
      revised.length,
      (i, j) => original[i] === revised[j]
    );
    expect(matches).toHaveLength(4);
    for (const [i, j] of matches) expect(original[i]).toBe(revised[j]);
    expect(
      matches.every(
        ([i, j], at) =>
          at === 0 || (i > matches[at - 1][0] && j > matches[at - 1][1])
      )
    ).toBe(true);
  });

  it("gives up past its budget and matches only the shared head and tail", () => {
    const size = 20_000;
    const started = performance.now();
    const matches = commonSubsequence(
      size + 2,
      size + 2,
      (i, j) =>
        (i === 0 && j === 0) ||
        (i === size + 1 && j === size + 1) ||
        (i > 0 && i <= size && j > 0 && j <= size && i === size + 1 - j)
    );
    expect(performance.now() - started).toBeLessThan(500);
    expect(matches).toEqual([
      [0, 0],
      [size + 1, size + 1],
    ]);
  });
});

describe("alignSequences", () => {
  it("keeps every element of two equal sequences", () => {
    expect(align([item("a"), item("b")], [item("a"), item("b")])).toEqual([
      { kind: "kept", original: 0, revised: 0 },
      { kind: "kept", original: 1, revised: 1 },
    ]);
  });

  it("pairs a lone element left on each side between two kept ones as a text change", () => {
    expect(
      align(
        [item("a"), item("b"), item("c")],
        [item("a"), item("x"), item("c")]
      )
    ).toEqual([
      { kind: "kept", original: 0, revised: 0 },
      { kind: "changed", original: 1, revised: 1, change: "text" },
      { kind: "kept", original: 2, revised: 2 },
    ]);
  });

  it("pairs elements that read the same but are written differently as a formatting change", () => {
    expect(align([item("a", "t")], [item("b", "t")])).toEqual([
      { kind: "changed", original: 0, revised: 0, change: "formatting" },
    ]);
  });

  it("prefers an element written the same to one that only reads the same", () => {
    expect(align([item("a", "t")], [item("b", "t"), item("a", "t")])).toEqual([
      { kind: "added", revised: 0 },
      { kind: "kept", original: 0, revised: 1 },
    ]);
  });

  it("does not pair elements of different kinds", () => {
    expect(align([item("a", "a", "p")], [item("b", "b", "table")])).toEqual([
      { kind: "removed", original: 0 },
      { kind: "added", revised: 0 },
    ]);
  });

  it("pairs two neighbouring edits in order", () => {
    expect(
      align(
        [item("a"), item("b"), item("z")],
        [item("x"), item("y"), item("z")]
      )
    ).toEqual([
      { kind: "changed", original: 0, revised: 0, change: "text" },
      { kind: "changed", original: 1, revised: 1, change: "text" },
      { kind: "kept", original: 2, revised: 2 },
    ]);
  });

  it("pairs as many as the shorter side holds and reports the rest as removed", () => {
    expect(
      align(
        [item("a"), item("b"), item("c"), item("z")],
        [item("x"), item("y"), item("z")]
      )
    ).toEqual([
      { kind: "changed", original: 0, revised: 0, change: "text" },
      { kind: "changed", original: 1, revised: 1, change: "text" },
      { kind: "removed", original: 2 },
      { kind: "kept", original: 3, revised: 2 },
    ]);
  });

  it("pairs past an element of another kind rather than stopping at it", () => {
    expect(
      align(
        [item("a"), item("b", "b", "table"), item("c")],
        [item("x"), item("y"), item("w")]
      )
    ).toEqual([
      { kind: "changed", original: 0, revised: 0, change: "text" },
      { kind: "removed", original: 1 },
      { kind: "changed", original: 2, revised: 1, change: "text" },
      { kind: "added", revised: 2 },
    ]);
  });

  it("pairs a table with the table after a removed paragraph", () => {
    expect(
      align([item("p"), item("t", "t", "table")], [item("u", "u", "table")])
    ).toEqual([
      { kind: "removed", original: 0 },
      { kind: "changed", original: 1, revised: 0, change: "text" },
    ]);
  });

  it("aligns around an insertion in the middle of a long run", () => {
    const original = ["a", "b", "c", "d"].map((id) => item(id));
    const revised = ["a", "b", "n", "c", "d"].map((id) => item(id));
    expect(align(original, revised)).toEqual([
      { kind: "kept", original: 0, revised: 0 },
      { kind: "kept", original: 1, revised: 1 },
      { kind: "added", revised: 2 },
      { kind: "kept", original: 2, revised: 3 },
      { kind: "kept", original: 3, revised: 4 },
    ]);
  });

  it("handles empty sequences", () => {
    expect(align([], [])).toEqual([]);
    expect(align([item("a")], [])).toEqual([{ kind: "removed", original: 0 }]);
  });

  it("aligns two long sequences written differently throughout without a table of both", () => {
    const size = 20_000;
    const original = Array.from({ length: size }, (_, at) =>
      item(`a${at}`, `${at}`)
    );
    const revised = Array.from({ length: size }, (_, at) =>
      item(`b${at}`, `${size - at}`)
    );
    const started = performance.now();
    const alignment = align(original, revised);
    expect(performance.now() - started).toBeLessThan(1_000);
    expect(alignment.filter((step) => step.kind === "changed")).toHaveLength(
      size
    );
  });
});
