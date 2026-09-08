/**
 * The lines a table's cells draw because of where they sit in the grid (`table/gridBorders`),
 * derived again once the grid or the border inputs of the table moved.
 */

import { tableStyleAttrs } from "../../docx/formatting";
import { parsePropsXml } from "../../ooxml/props";
import { cellFixes, sameFormattingInputs } from "../../table/gridBorders";
import type { DocumentDeriver } from "./displayDerivation";

export const tableDisplay: DocumentDeriver = {
  name: "table",
  nodeTypes: ["table"],
  derive(table, pos, _doc, context, previous) {
    if (previous !== null && sameFormattingInputs(previous, table)) return [];
    const tblPr: unknown = table.attrs.tblPr;
    const attrs = {
      ...table.attrs,
      ...tableStyleAttrs(
        typeof tblPr === "string" ? parsePropsXml(tblPr) : null,
        context.document.formatting
      ),
    };
    const current = table.hasMarkup(table.type, attrs, table.marks)
      ? table
      : table.type.create(attrs, table.content, table.marks);
    return [
      { pos, attrs },
      ...cellFixes(current).map((fix) => ({
        pos: pos + 1 + fix.pos,
        attrs: fix.attrs,
      })),
    ];
  },
};
