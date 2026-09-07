/**
 * The values a paragraph draws with, read out of the document's styles along the same path the
 * import takes (`docx/formatting`), so that a paragraph an edit built, one it re-aligned and one
 * the file arrived with all carry the same values for the same properties.
 */

import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";
import {
  type ParagraphPlacement,
  paragraphAttrsOf,
  resolveParagraph,
  resolveRun,
} from "../../docx/formatting";
import { docxSchema } from "../../schema";
import type { DisplayAttrs } from "../../schema/displayDerivation";
import { cellDefaultsAt, tableCellSources } from "../../table/gridBorders";
import { paragraphPPr } from "../paragraphEdits";
import type { DocumentDeriver } from "./displayDerivation";

/** Whether the paragraph carries no values at all, which is one the styles were never read into */
function carriesNothing(node: PMNode): boolean {
  return node.attrs.format === null && node.attrs.styleRun === null;
}

/**
 * Whether the paragraph can be left as it stands: it is the paragraph it was, its own properties
 * are as they were, and it carries values. A paragraph a composition held back carries none, so it
 * is read into whatever it maps back to.
 */
function stands(node: PMNode, previous: PMNode | null): boolean {
  return (
    previous !== null &&
    previous.attrs.pPr === node.attrs.pPr &&
    !carriesNothing(node)
  );
}

/** Whether the spot being composed in falls inside this paragraph */
function holdsComposition(
  node: PMNode,
  pos: number,
  composing: number | null
): boolean {
  return (
    composing !== null && composing >= pos && composing <= pos + node.nodeSize
  );
}

/**
 * The part of a table the paragraph stands in, which its table style dresses it by. null for a
 * paragraph of the body, and for one in a cell of a table this editor did not build.
 *
 * The cell is found by walking out of the paragraph rather than passed down, because the walk that
 * runs the derivers goes over the blocks and hands each of them nothing but its own position.
 */
function placementOf(doc: PMNode, pos: number): ParagraphPlacement | null {
  const $pos = doc.resolve(pos);
  for (let depth = $pos.depth; depth >= 2; depth -= 1) {
    if ($pos.node(depth).type !== docxSchema.nodes.tableCell) continue;
    const table = $pos.node(depth - 2);
    if (table.type !== docxSchema.nodes.table) return null;
    // A cell of the table's own map is named by where it stands inside the table's content
    const cell = $pos.before(depth) - $pos.start(depth - 2);
    return cellDefaultsAt(TableMap.get(table), cell, tableCellSources(table))
      .placement;
  }
  return null;
}

export const paragraphDisplay: DocumentDeriver = {
  name: "paragraph",
  nodeTypes: ["paragraph"],
  derive(node, pos, doc, context, previous) {
    if (
      stands(node, previous) ||
      holdsComposition(node, pos, context.composingAt)
    ) {
      return [];
    }
    const formatting = context.document.formatting;
    const paragraph = resolveParagraph(
      paragraphPPr(node),
      formatting,
      placementOf(doc, pos)
    );
    const derived: DisplayAttrs[] = [
      { pos, attrs: paragraphAttrsOf(paragraph) },
    ];
    node.forEach((child, offset) => {
      const mark = docxSchema.marks.run.isInSet(child.marks);
      if (!mark) return;
      const rPr: unknown = mark.attrs.rPr;
      derived.push({
        pos: pos + 1 + offset,
        mark: { type: mark.type, to: pos + 1 + offset + child.nodeSize },
        attrs: {
          format: resolveRun(
            typeof rPr === "string" ? rPr : null,
            paragraph,
            formatting
          ),
        },
      });
    });
    return derived;
  },
};
