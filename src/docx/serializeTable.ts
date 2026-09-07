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
import { elementXml, openTagXml, type XmlAttr } from "../ooxml/element";
import { DocxExportError } from "../ooxml/errors";
import { wName } from "../ooxml/names";
import {
  innerXml,
  type Props,
  parseProps,
  propsChild,
  renderProps,
  setChild,
} from "../ooxml/props";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import {
  preservedXml,
  rawAttrsOf,
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

/**
 * An element whose attributes the model decides, keeping whatever stood inside the one it replaces.
 *
 * These elements carry attributes alone in the schema, so what a producer put inside is nothing
 * this package models. It is still content the file arrived with, and a rebuilt block that quietly
 * dropped it would be a block the verifier (`./storyProjection`) cannot see all of.
 */
function modelled(
  name: string,
  attrs: readonly XmlAttr[],
  replacing: string | undefined
): string {
  const inner = replacing === undefined ? "" : innerXml(replacing);
  return elementXml(wName(name), attrs, inner === "" ? [] : [inner]);
}

/** A width that carries no number goes out with the 0 Word writes in its place */
function widthXml(
  name: string,
  width: TableWidth,
  replacing: string | undefined
): string {
  return modelled(
    name,
    [
      [wName("w"), `${widthNumber(width) ?? 0}`],
      [wName("type"), width.type],
    ],
    replacing
  );
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
  width: TableWidth | null
): Props {
  if (!width) return props;
  const replacing = propsChild(props.children, name)?.xml;
  return setChild(props, name, widthXml(name, width, replacing));
}

function tablePropsXml(table: PMNode): string {
  const props = propsOf(table.attrs.tblPr, "w:tblPr");
  const width = toTableWidth(table.attrs.tblW);
  // CT_Tbl requires tblPr even when the table has no properties (ECMA-376 Part 1, Annex A.1).
  return (
    renderProps(withWidth(props, "tblW", width)) ||
    elementXml(wName("tblPr"), [])
  );
}

/**
 * The grid, built fresh from the column widths the model holds, closed by the revision markup the
 * file arrived with.
 *
 * CT_TblGrid is `gridCol*` followed by an optional `tblGridChange` (ECMA-376 Part 1 17.4.48), so
 * the record of an earlier grid goes back where it stood however many columns now precede it. It
 * says what the grid was before it was last revised, which a column edit here does not change.
 */
function tableGridXml(table: PMNode): string {
  const gridCols = toGridCols(table.attrs.gridCols);
  const gridChange =
    typeof table.attrs.gridChange === "string" ? table.attrs.gridChange : "";
  if (gridCols.length === 0 && gridChange === "") return "";
  const cols = gridCols
    .map((w) => elementXml(wName("gridCol"), [[wName("w"), `${w}`]]))
    .join("");
  return elementXml(wName("tblGrid"), [], [cols + gridChange]);
}

type CellRole = "start" | "continue";

function cellPropsXml(cell: PMNode, role: CellRole): string {
  const props = propsOf(cell.attrs.tcPr, "w:tcPr");
  const colspan = spanCount(cell.attrs.colspan);
  const rowspan = spanCount(cell.attrs.rowspan);

  const gridSpan =
    colspan > 1
      ? modelled(
          "gridSpan",
          [[wName("val"), `${colspan}`]],
          propsChild(props.children, "gridSpan")?.xml
        )
      : null;
  // A continuing cell is written from the starting cell's properties, so it has nothing of its
  // own to keep, and keeping the starting cell's would copy it down the whole merge
  const vMerge =
    role === "continue"
      ? elementXml(wName("vMerge"), [])
      : rowspan > 1
        ? modelled(
            "vMerge",
            [[wName("val"), "restart"]],
            propsChild(props.children, "vMerge")?.xml
          )
        : null;

  const spanned = setChild(props, "gridSpan", gridSpan);
  const merged = setChild(spanned, "vMerge", vMerge);
  const width = toTableWidth(cell.attrs.tcW);
  return renderProps(withWidth(merged, "tcW", width));
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
      ? elementXml(wName("p"), [])
      : cell.children.map((block) => cellBlockXml(block, refs)).join("");
  const xml =
    openTagXml(wName("tc"), rawAttrsOf(cell.attrs.tcAttrs)) +
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
    openTagXml(wName("tr"), rawAttrsOf(row.attrs.trAttrs)) +
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
    openTagXml(wName("tbl"), rawAttrsOf(table.attrs.tblAttrs)) +
    tablePropsXml(table) +
    tableGridXml(table) +
    rows +
    "</w:tbl>"
  );
}
