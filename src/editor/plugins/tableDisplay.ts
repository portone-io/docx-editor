/**
 * The lines a table's cells draw because of where they sit in the grid (`table/gridBorders`),
 * derived again once the grid or the border inputs of the table moved.
 */

import { cellFixes, sameFormattingInputs } from "../../table/gridBorders";
import type { DocumentDeriver } from "./displayDerivation";

export const tableDisplay: DocumentDeriver = {
  name: "table",
  nodeTypes: ["table"],
  derive(table, pos, _doc, _context, previous) {
    if (previous !== null && sameFormattingInputs(previous, table)) return [];
    // A fix names the cell by its position within the table's content
    return cellFixes(table).map((fix) => ({
      pos: pos + 1 + fix.pos,
      attrs: fix.attrs,
    }));
  },
};
