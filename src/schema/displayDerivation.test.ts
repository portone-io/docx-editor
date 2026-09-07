// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { DocAttrStep, type Step, Transform } from "prosemirror-transform";
import { describe, expect, it } from "vitest";
import { changesOnlyDisplayAttrs } from "./displayDerivation";
import { docxSchema } from "./index";

const paragraph = (attrs: Record<string, unknown> = {}) =>
  docxSchema.nodes.paragraph.create(attrs, [docxSchema.text("text")]);

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
