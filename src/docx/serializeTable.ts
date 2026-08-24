/**
 * Rebuilds an edited table back into OOXML.
 *
 * The editor model holds a vertical merge as a single `rowspan` on the starting cell, but docx has to
 * write an empty cell out on every continuing row. Those empty cells are recreated here.
 *
 * A cell's `<w:tcPr>` is taken from the original with only the merge counts and the width swapped in.
 * That way, when a horizontal merge shrinks from two cells to one, the `w:gridSpan` goes away with it.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  spanCount,
  type TableWidth,
  toGridCols,
  toTableWidth,
  widthNumber,
} from "../model/format";
import { DocxExportError } from "../ooxml/errors";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import {
  type Props,
  parseProps,
  renderProps,
  setPropsChild,
  TBL_PR_ORDER,
  TC_PR_ORDER,
} from "./propsXml";
import {
  openTag,
  preservedXml,
  serializeParagraph,
} from "./serializeParagraph";

/** Splits the original formatting fragment up by child. With no fragment, we start from an empty one */
function propsOf(xml: unknown, tag: string): Props {
  if (typeof xml !== "string") return { tag, attrs: null, children: [] };
  const parsed = parseProps(xml);
  // We stop rather than quietly throwing away formatting we could not make out
  if (!parsed) {
    throw new DocxExportError("malformed-xml", `${tag} cannot be rewritten`);
  }
  return parsed;
}

/** A width that carries no number goes out with the 0 Word writes in its place */
function widthXml(name: string, width: TableWidth): string {
  return `<w:${name} w:w="${widthNumber(width) ?? 0}" w:type="${width.type}"/>`;
}

/**
 * Rewrites the child that records the width with the value the model holds.
 *
 * If the model holds no width (that is, a width we failed to read), the original is left untouched.
 * That way a width whose meaning we do not know survives editing unchanged.
 */
function withWidth(
  props: Props,
  name: "tblW" | "tcW",
  width: TableWidth | null,
  order: readonly string[]
): Props {
  if (!width) return props;
  return {
    ...props,
    children: setPropsChild(props.children, name, widthXml(name, width), order),
  };
}

function tablePropsXml(table: PMNode): string {
  const props = propsOf(table.attrs.tblPr, "w:tblPr");
  const width = toTableWidth(table.attrs.tblW);
  return renderProps(withWidth(props, "tblW", width, TBL_PR_ORDER));
}

/** The grid is always built fresh from the column widths the model holds. Revision markup does not come along */
function tableGridXml(table: PMNode): string {
  const gridCols = toGridCols(table.attrs.gridCols);
  if (gridCols.length === 0) return "";
  const cols = gridCols.map((w) => `<w:gridCol w:w="${w}"/>`).join("");
  return `<w:tblGrid>${cols}</w:tblGrid>`;
}

type CellRole = "start" | "continue";

function cellPropsXml(cell: PMNode, role: CellRole): string {
  const props = propsOf(cell.attrs.tcPr, "w:tcPr");
  const colspan = spanCount(cell.attrs.colspan);
  const rowspan = spanCount(cell.attrs.rowspan);

  const gridSpan = colspan > 1 ? `<w:gridSpan w:val="${colspan}"/>` : null;
  const vMerge =
    role === "continue"
      ? "<w:vMerge/>"
      : rowspan > 1
        ? '<w:vMerge w:val="restart"/>'
        : null;

  let children = setPropsChild(
    props.children,
    "gridSpan",
    gridSpan,
    TC_PR_ORDER
  );
  children = setPropsChild(children, "vMerge", vMerge, TC_PR_ORDER);
  const width = toTableWidth(cell.attrs.tcW);
  return renderProps(
    withWidth({ ...props, children }, "tcW", width, TC_PR_ORDER)
  );
}

function cellBlockXml(block: PMNode, refs: ExportRefs): string {
  if (block.type.name === "paragraph") return serializeParagraph(block, refs);
  if (block.type.name === "rawBlock") return preservedXml(block);
  if (block.type.name === "table") return serializeTable(block, refs);
  throw new DocxExportError(
    "unsupported-content",
    `block that cannot go inside a table cell: ${block.type.name}`
  );
}

/**
 * Puts the content control that wrapped this cell in the original back around it.
 *
 * Only a cell that starts where it stands carries the wrapper. The empty cells rebuilt on the
 * continuing rows of a vertical merge are new cells, so wrapping them would duplicate the control.
 */
function wrapInSdt(xml: string, cell: PMNode, role: CellRole): string {
  const prefix: unknown = cell.attrs.sdtPrefix;
  if (role !== "start" || typeof prefix !== "string") return xml;
  return `${prefix}<w:sdtContent>${xml}</w:sdtContent></w:sdt>`;
}

function cellXml(cell: PMNode, role: CellRole, refs: ExportRefs): string {
  // A continuing cell only holds the spot, so it carries a single empty paragraph
  const body =
    role === "continue"
      ? "<w:p/>"
      : cell.children.map((block) => cellBlockXml(block, refs)).join("");
  const xml =
    openTag("w:tc", cell.attrs.tcAttrs) +
    cellPropsXml(cell, role) +
    body +
    "</w:tc>";
  return wrapInSdt(xml, cell, role);
}

interface Placed {
  cell: PMNode;
  role: CellRole;
}

/** A merged cell that continues into the rows below */
interface Covering {
  column: number;
  colspan: number;
  cell: PMNode;
  rowsLeft: number;
}

/**
 * Fills in the cells to be placed on one row, starting from the left.
 * Where a merged cell has come down from above, a continuing cell goes into that spot.
 */
function placeRow(row: PMNode, covering: Covering[]): Placed[] {
  const placed: Placed[] = [];
  let column = 0;
  let next = 0;

  for (;;) {
    const cover = covering.find(
      (entry) => entry.column === column && entry.rowsLeft > 0
    );
    if (cover) {
      placed.push({ cell: cover.cell, role: "continue" });
      cover.rowsLeft -= 1;
      column += cover.colspan;
      continue;
    }
    if (next >= row.childCount) break;
    const cell = row.child(next);
    next += 1;
    const colspan = spanCount(cell.attrs.colspan);
    const rowspan = spanCount(cell.attrs.rowspan);
    placed.push({ cell, role: "start" });
    if (rowspan > 1) {
      covering.push({ column, colspan, cell, rowsLeft: rowspan - 1 });
    }
    column += colspan;
  }
  return placed;
}

function rowXml(row: PMNode, covering: Covering[], refs: ExportRefs): string {
  const cells = placeRow(row, covering)
    .map((placed) => cellXml(placed.cell, placed.role, refs))
    .join("");
  const tblPrEx: unknown = row.attrs.tblPrEx;
  const trPr: unknown = row.attrs.trPr;
  // Inside a row the table property exceptions come ahead of the row properties
  return (
    openTag("w:tr", row.attrs.trAttrs) +
    (typeof tblPrEx === "string" ? tblPrEx : "") +
    (typeof trPr === "string" ? trPr : "") +
    cells +
    "</w:tr>"
  );
}

export function serializeTable(
  table: PMNode,
  refs: ExportRefs = NO_EXPORT_REFS
): string {
  const covering: Covering[] = [];
  const rows = table.children
    .map((row) => rowXml(row, covering, refs))
    .join("");
  // A merge left over unconsumed means the grid is out of alignment, so we do not let it pass quietly
  if (covering.some((entry) => entry.rowsLeft > 0)) {
    throw new DocxExportError(
      "invalid-table",
      "a vertical merge in the table reaches past the last row"
    );
  }
  return (
    openTag("w:tbl", table.attrs.tblAttrs) +
    tablePropsXml(table) +
    tableGridXml(table) +
    rows +
    "</w:tbl>"
  );
}
