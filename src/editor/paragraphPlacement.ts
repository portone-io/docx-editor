/** Finds the table-style conditions at an editor position for both reads and writes. */
import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";
import type { ParagraphPlacement } from "../docx/formatting";
import { docxSchema } from "../schema";
import { cellDefaultsAt, tableCellSources } from "../table/gridBorders";

/**
 * The part of a table the paragraph stands in, which its table style dresses it by. null for a
 * paragraph of the body, and for one in a cell of a table this editor did not build.
 *
 * The cell is found by walking out of the paragraph rather than passed down, because the walk that
 * runs the derivers goes over the blocks and hands each of them nothing but its own position.
 */
export function paragraphPlacementAt(
  doc: PMNode,
  pos: number
): ParagraphPlacement | null {
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
