/**
 * The fake table documents shared by the table manipulation tests, built in memory
 * without any docx file.
 *
 * The schema borrows only the node names and attr names of the document model and does
 * not interpret any values.
 * That way the tests also confirm that table manipulation needs to know nothing about
 * the view or the file layer.
 */

import { type Node as PMNode, Schema } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { TableWidth } from "../../model/format";

export const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { group: "block", content: "inline*", toDOM: () => ["p", 0] },
    text: { group: "inline" },
    table: {
      group: "block",
      content: "tableRow+",
      tableRole: "table",
      isolating: true,
      attrs: {
        srcId: { default: null },
        tblAttrs: { default: null },
        tblPr: { default: null },
        tblW: { default: null },
        gridCols: { default: null },
        format: { default: null },
        styleInside: { default: null },
        styleCellMargins: { default: null },
      },
      toDOM: () => ["table", ["tbody", 0]],
    },
    tableRow: {
      content: "tableCell+",
      tableRole: "row",
      attrs: {
        trAttrs: { default: null },
        tblPrEx: { default: null },
        trPr: { default: null },
        format: { default: null },
      },
      toDOM: () => ["tr", 0],
    },
    tableCell: {
      content: "block+",
      tableRole: "cell",
      isolating: true,
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        colwidth: { default: null },
        tcAttrs: { default: null },
        tcPr: { default: null },
        tcW: { default: null },
        format: { default: null },
        sdtPrefix: { default: null },
        sdtContentsLocked: { default: false },
        sdtDeletionLocked: { default: false },
      },
      toDOM: () => ["td", 0],
    },
  },
});

export interface CellAttrs {
  colspan?: number;
  rowspan?: number;
  colwidth?: number[] | null;
  tcAttrs?: string | null;
  tcPr?: string | null;
  tcW?: TableWidth | null;
  format?: Record<string, unknown> | null;
  sdtPrefix?: string | null;
  sdtContentsLocked?: boolean;
  sdtDeletionLocked?: boolean;
}

export const dxa = (twips: number): TableWidth => ({ type: "dxa", twips });
export const pct = (fiftieths: number): TableWidth => ({
  type: "pct",
  fiftieths,
});

export function cell(text: string, attrs: CellAttrs = {}): PMNode {
  return schema.nodes.tableCell.create(
    {
      colspan: 1,
      rowspan: 1,
      colwidth: null,
      tcAttrs: null,
      tcPr: null,
      tcW: null,
      format: null,
      sdtPrefix: null,
      sdtContentsLocked: false,
      sdtDeletionLocked: false,
      ...attrs,
    },
    schema.nodes.paragraph.create(null, text ? schema.text(text) : undefined)
  );
}

export const row = (...cells: PMNode[]) =>
  schema.nodes.tableRow.create(null, cells);

export interface RowAttrs {
  trAttrs?: string | null;
  tblPrEx?: string | null;
  trPr?: string | null;
  format?: Record<string, unknown> | null;
}

export const rowWith = (attrs: RowAttrs, ...cells: PMNode[]) =>
  schema.nodes.tableRow.create(
    { trAttrs: null, tblPrEx: null, trPr: null, format: null, ...attrs },
    cells
  );

export function tableDoc(
  rows: PMNode[],
  tableAttrs: Record<string, unknown> | null = null
): PMNode {
  return schema.nodes.doc.create(null, [
    schema.nodes.table.create(tableAttrs, rows),
  ]);
}

/** A contract-style table mixing a rowspan-3 "PartyB" cell with horizontal merges */
export function mergedContractDoc(): PMNode {
  return tableDoc([
    row(
      cell("PartyB", { rowspan: 3, tcW: dxa(1000) }),
      cell("Company", { tcW: dxa(1500) }),
      cell("Acme Corp", { colspan: 2, tcW: dxa(4000) })
    ),
    row(
      cell("Address", { tcW: dxa(1500) }),
      cell("10 Main Street", { colspan: 2, tcW: dxa(4000) })
    ),
    row(
      cell("Contact", { tcW: dxa(1500) }),
      cell("555-0123", { tcW: dxa(2000) }),
      cell("CEO Jane Doe", { tcW: dxa(2000) })
    ),
  ]);
}

/** A 3x3 table with no merges */
export function uniformDoc(): PMNode {
  const mk = (a: string, b: string, c: string) =>
    row(
      cell(a, { tcW: dxa(3000) }),
      cell(b, { tcW: dxa(3000) }),
      cell(c, { tcW: dxa(3000) })
    );
  return tableDoc([
    mk("Item", "Detail", "Amount"),
    mk("Design", "Requirements", "3,000,000"),
    mk("Build", "Implementation", "7,000,000"),
  ]);
}

/** A table with two overlapping rowspans, as in a real contract */
export function sampleContractDoc(): PMNode {
  return tableDoc([
    row(
      cell("Parties", { rowspan: 5, tcW: dxa(500) }),
      cell("Client", { rowspan: 4, tcW: dxa(1200) }),
      cell("Company", { tcW: dxa(1200) }),
      cell("Acme Corp", { colspan: 2, tcW: dxa(4000) })
    ),
    row(
      cell("Address", { tcW: dxa(1200) }),
      cell("10 Main Street", { colspan: 2, tcW: dxa(4000) })
    ),
    row(
      cell("Contact", { tcW: dxa(1200) }),
      cell("555-0100", { colspan: 2, tcW: dxa(4000) })
    ),
    row(
      cell("Manager", { tcW: dxa(1200) }),
      cell("Phone", { colspan: 2, tcW: dxa(4000) })
    ),
    row(
      cell("Instructor", { tcW: dxa(1200) }),
      cell("Company", { tcW: dxa(1200) }),
      cell("[  ]", { colspan: 2, tcW: dxa(4000) })
    ),
  ]);
}

export function firstTable(doc: PMNode): { table: PMNode; start: number } {
  const found: { table: PMNode; start: number }[] = [];
  doc.descendants((node, pos) => {
    if (found.length > 0) return false;
    if (node.type.spec.tableRole === "table") {
      found.push({ table: node, start: pos + 1 });
      return false;
    }
    return true;
  });
  const first = found[0];
  if (!first) throw new Error("no table");
  return first;
}

export function posOfText(doc: PMNode, needle: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found < 0 && node.isText && node.text?.includes(needle))
      found = pos + 1;
  });
  if (found < 0) throw new Error(`text not found: ${needle}`);
  return found;
}

/**
 * The position of the cell itself that holds the given text. Used when building a cell
 * selection
 */
export function cellPosOfText(doc: PMNode, needle: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found >= 0) return false;
    if (
      node.type.spec.tableRole === "cell" &&
      node.textContent.includes(needle)
    ) {
      found = pos;
      return false;
    }
    return true;
  });
  if (found < 0) throw new Error(`cell not found: ${needle}`);
  return found;
}

/** Builds a state with the cursor placed in the cell holding the given text */
export function stateWithCursorIn(doc: PMNode, needle: string): EditorState {
  const state = EditorState.create({ doc, schema });
  const selection = TextSelection.near(doc.resolve(posOfText(doc, needle)));
  return state.apply(state.tr.setSelection(selection));
}

/** Builds a cell selection with those two cells as its corners */
export function stateWithCellsSelected(
  doc: PMNode,
  anchorText: string,
  headText: string
): EditorState {
  const state = EditorState.create({ doc, schema });
  const selection = CellSelection.create(
    doc,
    cellPosOfText(doc, anchorText),
    cellPosOfText(doc, headText)
  );
  return state.apply(state.tr.setSelection(selection));
}

export function cellWithText(table: PMNode, text: string): PMNode | undefined {
  let found: PMNode | undefined;
  table.descendants((node) => {
    if (
      !found &&
      node.type.spec.tableRole === "cell" &&
      node.textContent === text
    )
      found = node;
  });
  return found;
}

export function rowsOf(table: PMNode): PMNode[] {
  const rows: PMNode[] = [];
  table.forEach((current) => rows.push(current));
  return rows;
}

export function cellsOf(rowNode: PMNode): PMNode[] {
  const cells: PMNode[] = [];
  rowNode.forEach((current) => cells.push(current));
  return cells;
}
