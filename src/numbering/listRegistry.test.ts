// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { listsWorn, newListsOf, newListsValue } from "./listRegistry";
import { templateList } from "./listTemplate";
import type { NewList } from "./parseNumbering";

function register(...numIds: readonly number[]): Map<number, NewList> {
  return new Map(
    numIds.map((numId) => [
      numId,
      templateList(numId % 2 === 0 ? "numbered" : "bullet"),
    ])
  );
}

/** The value a document node would carry for these lists */
function carried(...numIds: readonly number[]): unknown {
  return newListsValue(register(...numIds));
}

describe("what the document node carries", () => {
  it.each([
    ["start", 1e21],
    ["ilvl", 9],
    ["restartAfterLevel", "0"],
    ["legal", "false"],
    ["run", { bold: true }],
    ["tabStops", [{ positionPt: 36, alignment: "left" }]],
    [
      "indent",
      {
        startTwips: 1.5,
        endTwips: null,
        hangingTwips: null,
        firstLineTwips: null,
      },
    ],
    [
      "indent",
      {
        startTwips: null,
        endTwips: null,
        hangingTwips: -1,
        firstLineTwips: null,
      },
    ],
  ])(
    "refuses a definition whose %s cannot be written faithfully (%j)",
    (field, replacement) => {
      const value = carried(2);
      if (!Array.isArray(value)) throw new Error("expected a registry value");
      value[0].levels[0][field] = replacement;
      expect(newListsOf(value).size).toBe(0);
    }
  );

  it("refuses duplicate list ids and duplicate levels instead of losing one definition", () => {
    const value = carried(2);
    if (!Array.isArray(value)) throw new Error("expected a registry value");
    expect(newListsOf([...value, ...value]).size).toBe(0);
    value[0].levels.push(value[0].levels[0]);
    expect(newListsOf(value).size).toBe(0);
  });
  it("reads back every definition it was given", () => {
    const lists = register(2, 5);

    expect(newListsOf(newListsValue(lists))).toEqual(lists);
  });

  it("carries nothing for a document that started no list", () => {
    expect(newListsValue(new Map())).toBeNull();
    expect(newListsOf(null).size).toBe(0);
  });

  it("is the same value however the lists were registered", () => {
    const forwards = new Map([...register(2), ...register(5)]);
    const backwards = new Map([...register(5), ...register(2)]);

    expect(newListsValue(forwards)).toEqual(newListsValue(backwards));
  });

  it("reads no register at all out of an entry that is not a definition", () => {
    expect(newListsOf("not a register").size).toBe(0);
    expect(newListsOf([{ numId: 2 }]).size).toBe(0);
    expect(newListsOf([{ numId: 2, levels: [] }]).size).toBe(0);
  });

  it("turns down a level counting in a format nothing can spell", () => {
    const value = carried(2);
    const levels = Array.isArray(value) ? value[0].levels : [];
    levels[0].format = "cardinalText";

    expect(newListsOf(value).size).toBe(0);
  });

  it("turns down a level that lost a value the writer needs", () => {
    for (const field of ["text", "start", "suffix", "align", "ilvl"]) {
      const value = carried(2);
      const levels = Array.isArray(value) ? value[0].levels : [];
      delete levels[0][field];

      expect(newListsOf(value).size, `dropping ${field}`).toBe(0);
    }
  });
});

describe("the lists an edit leaves standing", () => {
  it("drops the definitions of the lists nothing wears any more", () => {
    const lists = register(2, 5);

    expect([...listsWorn(lists, new Set([5])).keys()]).toEqual([5]);
  });

  it("is the register itself when every list is still worn", () => {
    const lists = register(2, 5);

    expect(listsWorn(lists, new Set([2, 5, 9]))).toBe(lists);
  });
});
