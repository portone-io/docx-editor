// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { DocAttrStep, type Step, Transform } from "prosemirror-transform";
import { describe, expect, it } from "vitest";
import {
  changesOnlyDisplayAttrs,
  type DisplayAttrs,
  type DisplayDeriver,
  deriveDisplay,
} from "./displayDerivation";
import { docxSchema } from "./index";

const paragraph = (attrs: Record<string, unknown> = {}, text = "text") =>
  docxSchema.nodes.paragraph.create(attrs, [docxSchema.text(text)]);

/** A table of one cell, standing first in a document, so its cell begins at position 3 */
function tableDoc(tableAttrs: Record<string, unknown>): PMNode {
  return docxSchema.nodes.doc.create(null, [
    docxSchema.nodes.table.create(tableAttrs, [
      docxSchema.nodes.tableRow.create(null, [
        docxSchema.nodes.tableCell.create({ sdtContentsLocked: true }, [
          paragraph(),
        ]),
      ]),
    ]),
  ]);
}

const CELL_POS = 2;

/** The one step this edit of the document writes, and the document it was written against */
function stepOf(
  doc: PMNode,
  edit: (transform: Transform) => void
): { step: Step; doc: PMNode } {
  const transform = new Transform(doc);
  edit(transform);
  const step = transform.steps[0];
  if (step === undefined || transform.steps.length !== 1) {
    throw new Error("the edit was expected to write one step");
  }
  return { step, doc };
}

function rewriting(
  doc: PMNode,
  pos: number,
  attrs: Record<string, unknown>
): { step: Step; doc: PMNode } {
  const node = doc.nodeAt(pos);
  if (node === null) throw new Error(`no node at ${pos}`);
  return stepOf(doc, (transform) =>
    transform.setNodeMarkup(pos, null, { ...node.attrs, ...attrs })
  );
}

describe("a step that changes only display attrs", () => {
  const doc = docxSchema.nodes.doc.create(null, [
    paragraph({ pPr: "<w:pPr/>", format: null, styleRun: null }),
  ]);

  it("is a node rewritten where it stands with every source attr the same", () => {
    const { step } = rewriting(doc, 0, {
      format: { align: "center" },
      styleRun: { bold: true },
    });
    expect(changesOnlyDisplayAttrs(step, doc)).toBe(true);
  });

  it("is an attr step naming a display attr, and not one naming a source attr", () => {
    const display = stepOf(doc, (transform) =>
      transform.setNodeAttribute(0, "format", { align: "center" })
    );
    const source = stepOf(doc, (transform) =>
      transform.setNodeAttribute(0, "pPr", "<w:pPr><w:jc/></w:pPr>")
    );
    expect(changesOnlyDisplayAttrs(display.step, doc)).toBe(true);
    expect(changesOnlyDisplayAttrs(source.step, doc)).toBe(false);
  });

  it("is not a node rewritten with a source attr changed beside a display attr", () => {
    const { step } = rewriting(doc, 0, {
      pPr: '<w:pPr><w:jc w:val="center"/></w:pPr>',
      format: { align: "center" },
    });
    expect(changesOnlyDisplayAttrs(step, doc)).toBe(false);
  });

  /**
   * A re-derivation spreads the node's own attrs, so a source attr it leaves alone is the very
   * same value. A value rebuilt to read the same is reported as a change rather than hidden, and
   * the guards judge it as they would any edit.
   */
  it("is not a node rewritten with a source attr equal by value alone", () => {
    const widths = tableDoc({ gridCols: [1000, 1000] });
    const { step } = rewriting(widths, 0, {
      gridCols: [1000, 1000],
      format: { width: 1 },
    });
    expect(changesOnlyDisplayAttrs(step, widths)).toBe(false);
  });

  it("is not a node rewritten with a lock flag lifted, whichever way the step is written", () => {
    const locked = tableDoc({});
    const whole = rewriting(locked, CELL_POS, {
      sdtContentsLocked: false,
      format: { background: "#FFFFFF" },
    });
    const oneAttr = stepOf(locked, (transform) =>
      transform.setNodeAttribute(CELL_POS, "sdtContentsLocked", false)
    );
    expect(locked.nodeAt(CELL_POS)?.attrs.sdtContentsLocked).toBe(true);
    expect(changesOnlyDisplayAttrs(whole.step, locked)).toBe(false);
    expect(changesOnlyDisplayAttrs(oneAttr.step, locked)).toBe(false);
  });

  it("is not a step that puts content anywhere, nor one on the document itself", () => {
    const typed = stepOf(doc, (transform) =>
      transform.insert(1, docxSchema.text("x"))
    );
    const attributed = stepOf(doc, (transform) =>
      transform.step(new DocAttrStep("sectPr", null))
    );
    expect(changesOnlyDisplayAttrs(typed.step, doc)).toBe(false);
    expect(changesOnlyDisplayAttrs(attributed.step, doc)).toBe(false);
  });
});

/** A context the derivers below are handed and never read, standing for whatever a caller resolves against */
const CONTEXT = "the context";

type Deriver = DisplayDeriver<string>;

function deriverFor(
  nodeTypes: Deriver["nodeTypes"],
  derive: Deriver["derive"],
  name = "test"
): Deriver {
  return { name, nodeTypes, derive };
}

/** The paragraph at this spot rewritten with these attrs, which is what most derivers below answer */
const at = (pos: number, attrs: Record<string, unknown>): DisplayAttrs[] => [
  { pos, attrs },
];

const nobody = (): null => null;

function derived(
  doc: PMNode,
  derivers: readonly Deriver[],
  previousOf: (node: PMNode, pos: number) => PMNode | null = nobody
): Transform {
  const transform = new Transform(doc);
  deriveDisplay(transform, CONTEXT, derivers, previousOf);
  return transform;
}

const alignOf = (node: PMNode | null | undefined): unknown =>
  node?.attrs.format?.align;

describe("the walk that writes what the derivers hand back", () => {
  const twoParagraphs = () =>
    docxSchema.nodes.doc.create(null, [
      paragraph({ pPr: "<w:pPr/>" }, "first"),
      paragraph({}, "second"),
    ]);

  it("writes no step for a deriver that hands back the node's own attrs", () => {
    const doc = twoParagraphs();
    const transform = derived(doc, [
      deriverFor(["paragraph"], (node, pos) => at(pos, node.attrs)),
    ]);

    expect(transform.steps).toEqual([]);
    expect(transform.doc).toBe(doc);
  });

  it("writes one step per node whose display attrs change, and none for a node handed nothing", () => {
    const doc = twoParagraphs();
    const transform = derived(doc, [
      deriverFor(["paragraph"], (node, pos) =>
        node.textContent === "first"
          ? at(pos, { format: { align: "center" } })
          : []
      ),
    ]);

    expect(transform.steps).toHaveLength(1);
    expect(alignOf(transform.doc.child(0))).toBe("center");
    expect(transform.doc.child(1)).toBe(doc.child(1));
  });

  it("asks the derivers of one node in registration order, each reading what the one before wrote", () => {
    const first = deriverFor(
      ["paragraph"],
      (_node, pos) => at(pos, { format: { align: "left" } }),
      "first"
    );
    const second = deriverFor(
      ["paragraph"],
      (node, pos) =>
        at(pos, {
          format: { align: alignOf(node) === "left" ? "right" : "center" },
        }),
      "second"
    );

    expect(
      alignOf(derived(twoParagraphs(), [first, second]).doc.child(0))
    ).toBe("right");
    expect(
      alignOf(derived(twoParagraphs(), [second, first]).doc.child(0))
    ).toBe("left");
  });

  /** A deriver cannot rewrite a source attr, whatever it hands back, so the walk stays no edit */
  it("writes only the display attrs of what a deriver hands back", () => {
    const doc = twoParagraphs();
    const transform = derived(doc, [
      deriverFor(["paragraph"], (_node, pos) =>
        at(pos, {
          pPr: "<w:pPr><w:jc/></w:pPr>",
          srcId: "elsewhere:body:9",
          format: { align: "center" },
        })
      ),
    ]);

    const written = transform.doc.child(0);
    expect(alignOf(written)).toBe("center");
    expect(written.attrs.pPr).toBe("<w:pPr/>");
    expect(written.attrs.srcId).toBeNull();
  });

  const tabled = () =>
    docxSchema.nodes.doc.create(null, [
      paragraph({}, "body"),
      docxSchema.nodes.table.create(null, [
        docxSchema.nodes.tableRow.create(null, [
          docxSchema.nodes.tableCell.create(null, [paragraph({}, "left")]),
          docxSchema.nodes.tableCell.create(null, [paragraph({}, "right")]),
        ]),
      ]),
    ]);

  it("asks once for every node of a type it answers for, the ones inside a table included", () => {
    const asked: string[] = [];
    derived(tabled(), [
      deriverFor(["paragraph", "table"], (node) => {
        asked.push(node.type.name === "table" ? "table" : node.textContent);
        return [];
      }),
    ]);

    expect(asked).toEqual(["body", "table", "left", "right"]);
  });

  it("writes what a deriver hands back for a node inside the one it answers for", () => {
    const doc = tabled();
    const transform = derived(doc, [
      deriverFor(["table"], (table, pos) => {
        const cells: DisplayAttrs[] = [];
        table.descendants((node, offset) => {
          if (node.type.name === "tableCell") {
            cells.push({
              pos: pos + 1 + offset,
              attrs: { format: { background: "#FFFF00" } },
            });
          }
          return node.type.name !== "tableCell";
        });
        return cells;
      }),
    ]);

    const row = transform.doc.child(1).child(0);
    expect(transform.steps).toHaveLength(2);
    expect(row.child(0).attrs.format).toEqual({ background: "#FFFF00" });
    expect(row.child(1).attrs.format).toEqual({ background: "#FFFF00" });
    // Nothing inside the cells was rewritten
    expect(row.child(0).child(0)).toBe(doc.child(1).child(0).child(0).child(0));
  });

  it("hands each node what the caller says it maps back to", () => {
    const doc = twoParagraphs();
    const handed: (PMNode | null)[] = [];
    derived(
      doc,
      [
        deriverFor(["paragraph"], (_node, _pos, _doc, _context, previous) => {
          handed.push(previous);
          return [];
        }),
      ],
      (node) => (node.textContent === "first" ? node : null)
    );

    expect(handed).toEqual([doc.child(0), null]);
  });

  it("refuses a deriver answering for an inline type, which the walk over the blocks never reaches", () => {
    expect(() =>
      derived(twoParagraphs(), [deriverFor(["hardBreak"], () => [], "inline")])
    ).toThrow(/inline answers for hardBreak/);
  });

  it("names its node types by the schema", () => {
    const misspelt: Deriver = {
      name: "misspelt",
      // @ts-expect-error A name the schema does not have is refused when the deriver is written
      nodeTypes: ["tabel"],
      derive: () => [],
    };
    expect(() => derived(twoParagraphs(), [misspelt])).toThrow(/tabel/);
  });
});
