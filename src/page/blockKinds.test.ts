import type { Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { docxSchema } from "../schema";
import { type BlockKind, blockKindFor } from "./blockKinds";

function kind(name: string, matches: (node: PMNode) => boolean): BlockKind {
  return {
    name,
    matches,
    measure: () => ({
      candidates: [],
      minFirstPiece: 0,
      breakAfter: false,
      appliedHeight: 0,
    }),
    holdsCut: () => false,
    decorate: () => undefined,
  };
}

function paragraph(): PMNode {
  return docxSchema.nodes.paragraph.create({}, [docxSchema.text("text")]);
}

function table(): PMNode {
  return docxSchema.nodes.table.create({ gridCols: [1000] }, [
    docxSchema.nodes.tableRow.create({}, [
      docxSchema.nodes.tableCell.create({}, [paragraph()]),
    ]),
  ]);
}

describe("blockKindFor", () => {
  it("the first matching kind wins and the last kind matches every block", () => {
    const anything = kind("anything", () => true);
    const tables = kind(
      "tables",
      (node) => node.type.spec.tableRole === "table"
    );

    expect(blockKindFor([tables, anything], table()).name).toBe("tables");
    expect(blockKindFor([anything, tables], table()).name).toBe("anything");
    expect(blockKindFor([tables, anything], paragraph()).name).toBe("anything");

    // A registry whose last kind claims nothing leaves a block with no measurer at all, which
    // is a registration mistake rather than a block the engine may skip
    expect(() => blockKindFor([tables], paragraph())).toThrow(
      "no block kind matches a paragraph block"
    );
  });
});
