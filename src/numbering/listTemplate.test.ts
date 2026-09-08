// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { allocateList, listFor, templateList } from "./listTemplate";
import {
  EMPTY_NUMBERING,
  type Numbering,
  parseNumbering,
} from "./parseNumbering";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** A numbering part defining the lists these numbers name, one level each */
function defining(...numIds: readonly number[]): Numbering {
  return parseNumbering(
    `<w:numbering ${W_NS}>` +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
      '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
      "</w:abstractNum>" +
      numIds
        .map(
          (numId) =>
            `<w:num w:numId="${numId}"><w:abstractNumId w:val="0"/></w:num>`
        )
        .join("") +
      "</w:numbering>"
  );
}

describe("giving a new list a number", () => {
  it("uses an unused safe id when the highest id cannot be incremented", () => {
    const numbering = defining(1, Number.MAX_SAFE_INTEGER);
    const first = allocateList(numbering, [2], templateList("numbered"));
    const second = allocateList(first.numbering, [2], templateList("numbered"));
    expect([first.numId, second.numId]).toEqual([3, 4]);
  });
  it("takes the number after the highest one the document defines", () => {
    const { numId } = allocateList(defining(1, 4), [], templateList("bullet"));

    expect(numId).toBe(5);
  });

  it("passes over a number a paragraph wears that no definition stands behind", () => {
    const { numId } = allocateList(defining(1), [9], templateList("numbered"));

    expect(numId).toBe(10);
  });

  it("passes over the numbers already given to lists started before it", () => {
    const first = allocateList(defining(1), [], templateList("numbered"));
    const second = allocateList(first.numbering, [], templateList("numbered"));

    expect([first.numId, second.numId]).toEqual([2, 3]);
  });

  it("registers the definition it was handed under that number", () => {
    const list = templateList("bullet");
    const { numId, numbering } = allocateList(EMPTY_NUMBERING, [], list);

    expect(numbering.added.get(numId)).toBe(list);
    expect(listFor(numbering, numId)).toBe(list);
  });

  it("leaves the numbering it was given as it stands", () => {
    const numbering = defining(1);
    allocateList(numbering, [], templateList("bullet"));

    expect(numbering.added.size).toBe(0);
  });

  it("two lists of the same kind are defined the same way under different numbers", () => {
    const first = allocateList(defining(1), [], templateList("numbered"));
    const second = allocateList(first.numbering, [], templateList("numbered"));

    expect(first.numId).not.toBe(second.numId);
    expect(second.numbering.added.get(first.numId)).toEqual(
      second.numbering.added.get(second.numId)
    );
  });
});

describe("the definition standing behind a number", () => {
  it("prefers what the document wrote down to what was registered", () => {
    const defined = defining(1);
    const numbering: Numbering = {
      ...defined,
      added: new Map([[1, templateList("bullet")]]),
    };

    expect(listFor(numbering, 1)).toBe(defined.lists.get(1));
  });
});
