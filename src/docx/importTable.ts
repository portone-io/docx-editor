/**
 * Imports one OOXML table, converting vertical merge continuations to rowspans and layering style
 * values for display. Unsupported grids return null so the caller can preserve their source XML.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  type CellMargins,
  type TableFormat,
  toCellMargins,
  toInsideBorders,
} from "../model/format";
import { childValue, wAttr } from "../ooxml/units";
import {
  attrString,
  childByLocalName,
  elementChildren,
  serializeXml,
} from "../ooxml/xml";
import { docxSchema } from "../schema";
import {
  layerTableFormat,
  NO_STYLES,
  type StyleFormat,
  type StyleTable,
} from "./formatting";
import {
  buildParagraph,
  type ImportSources,
  NO_IMPORT_SOURCES,
} from "./importParagraph";
import { readSdtWrapper } from "./sdt";
import {
  type CellBorderDefaults,
  cellBorderDefaults,
  type GridSize,
  gridEdgesOf,
  type InsideBorders,
  layerCellMargins,
  layerInsideBorders,
  NO_CELL_MARGINS,
  NO_INSIDE_BORDERS,
  readCellFormat,
  readCellMarginsOf,
  readGridCols,
  readInsideBorders,
  readRowFormat,
  readTableFormat,
  readTableWidth,
  tblStyleIdOf,
} from "./tableFormatting";

type VerticalMerge = "restart" | "continue" | null;

/** The content control around a cell, as the cell carries it on (`sdtPrefix` in `schema`) */
interface CellControl {
  /** The `<w:sdt>` opening tag followed by everything that stood ahead of `<w:sdtContent>` */
  prefix: string;
  /** Whether it says the cell may not be edited */
  contentsLocked: boolean;
  /** Whether it says the control around the cell may not be deleted */
  deletionLocked: boolean;
}

interface RawCell {
  el: Element;
  gridSpan: number;
  vMerge: VerticalMerge;
  control: CellControl | null;
}

interface RawRow {
  el: Element;
  tblPrEx: Element | null;
  cells: RawCell[];
}

/** One cell to be built. Its rowspan grows as the continuing cells are counted */
interface CellDraft {
  el: Element;
  /** The row and the column the cell starts at, which is what tells it where its lines come from */
  row: number;
  col: number;
  colspan: number;
  rowspan: number;
  control: CellControl | null;
}

/** The horizontal merge count `w:gridSpan` states. One cell if it is absent */
function readGridSpan(tcPr: Element | null): number {
  const value = tcPr ? childValue(tcPr, "gridSpan") : null;
  const parsed = value === null ? null : Number.parseInt(value, 10);
  return parsed !== null && Number.isInteger(parsed) && parsed > 0 ? parsed : 1;
}

/** A `w:vMerge` whose val is restart starts a merge; with no val or with continue it is a continuing cell */
function readVerticalMerge(tcPr: Element | null): VerticalMerge {
  const vMerge = tcPr ? childByLocalName(tcPr, "vMerge") : null;
  if (!vMerge) return null;
  return wAttr(vMerge, "val") === "restart" ? "restart" : "continue";
}

/** A cell taken out of the content control that wrapped it, plus that control */
interface SdtCell {
  el: Element;
  control: CellControl;
}

/**
 * Takes the single cell out of a `w:sdt` content control.
 *
 * On top of the wrapper shape `readSdtWrapper` insists on, the content has to hold nothing but
 * one cell. Everything else is null, which leaves the whole table preserved.
 *
 * A cell that starts a vertical merge is a cell of the model in its own right, and `serializeTable`
 * puts its wrapper back around it. The cells that only continue the merge are created fresh on
 * export instead, and a wrapper of their own could not be recreated for them.
 */
function readSdtCell(el: Element): SdtCell | null {
  const wrapper = readSdtWrapper(el);
  if (!wrapper) return null;

  const inner = elementChildren(wrapper.content);
  const tc = inner[0];
  if (inner.length !== 1 || tc.localName !== "tc") return null;
  if (readVerticalMerge(childByLocalName(tc, "tcPr")) === "continue") {
    return null;
  }

  return {
    el: tc,
    control: {
      prefix: wrapper.prefix,
      contentsLocked: wrapper.contentsLocked,
      deletionLocked: wrapper.deletionLocked,
    },
  };
}

/**
 * Reads a single row. null if there is any child other than `w:tc`, `w:trPr` and `w:tblPrEx`,
 * or a content control we could not write back out.
 *
 * A `w:tblPrEx` states the table properties this one row departs from. We do not read it, so
 * the row is drawn with the table's own values, but it is carried along to go back out untouched.
 */
function readRow(el: Element): RawRow | null {
  const cells: RawCell[] = [];
  let tblPrEx: Element | null = null;
  for (const child of elementChildren(el)) {
    if (child.localName === "trPr") continue;
    if (child.localName === "tblPrEx") {
      tblPrEx = child;
      continue;
    }
    const sdt = child.localName === "sdt" ? readSdtCell(child) : null;
    if (!sdt && child.localName !== "tc") return null;
    const tc = sdt ? sdt.el : child;
    const tcPr = childByLocalName(tc, "tcPr");
    cells.push({
      el: tc,
      gridSpan: readGridSpan(tcPr),
      vMerge: readVerticalMerge(tcPr),
      control: sdt ? sdt.control : null,
    });
  }
  return cells.length > 0 ? { el, tblPrEx, cells } : null;
}

interface TableParts {
  tblPr: Element | null;
  tblGrid: Element | null;
  rows: RawRow[];
}

/** Splits a table into the three pieces it is made of. null if a child we do not know is mixed in */
function readTableParts(el: Element): TableParts | null {
  let tblPr: Element | null = null;
  let tblGrid: Element | null = null;
  const rows: RawRow[] = [];
  for (const child of elementChildren(el)) {
    if (child.localName === "tblPr") {
      tblPr = child;
      continue;
    }
    if (child.localName === "tblGrid") {
      tblGrid = child;
      continue;
    }
    if (child.localName !== "tr") return null;
    const row = readRow(child);
    if (!row) return null;
    rows.push(row);
  }
  return rows.length > 0 ? { tblPr, tblGrid, rows } : null;
}

/**
 * Counts the continuing cells and turns them into rowspans.
 * Returns, per row, the sequence of cells to build; null if the grid does not add up.
 */
function resolveVerticalMerges(
  rows: RawRow[],
  width: number
): CellDraft[][] | null {
  const drafts: CellDraft[][] = [];
  // starting column -> the merged cell currently being continued
  let open = new Map<number, CellDraft>();

  for (const [index, row] of rows.entries()) {
    const rowDrafts: CellDraft[] = [];
    const next = new Map<number, CellDraft>();
    let col = 0;

    for (const cell of row.cells) {
      if (cell.vMerge === "continue") {
        const chain = open.get(col);
        // With no cell to continue from, or with a different width, it cannot be moved onto a rectangular grid
        if (!chain || chain.colspan !== cell.gridSpan) return null;
        chain.rowspan += 1;
        next.set(col, chain);
      } else {
        const draft: CellDraft = {
          el: cell.el,
          row: index,
          col,
          colspan: cell.gridSpan,
          rowspan: 1,
          control: cell.control,
        };
        rowDrafts.push(draft);
        if (cell.vMerge === "restart") next.set(col, draft);
      }
      col += cell.gridSpan;
    }

    if (col !== width) return null;
    // A row with no cell to build at all cannot be held in the editor model
    if (rowDrafts.length === 0) return null;
    drafts.push(rowDrafts);
    open = next;
  }

  return drafts;
}

/** The sum, per row, of the columns its cells cover. null if it differs from row to row */
function gridWidthOf(rows: RawRow[]): number | null {
  const widths = rows.map((row) =>
    row.cells.reduce((sum, cell) => sum + cell.gridSpan, 0)
  );
  const first = widths[0];
  return widths.every((width) => width === first) ? first : null;
}

/** A single block inside a cell. If it is not a paragraph, or the paragraph cannot be modelled, it holds on to its original XML as is */
function buildCellBlock(el: Element, sources: ImportSources): PMNode {
  if (el.localName === "p") {
    const paragraph = buildParagraph(el, null, sources);
    if (paragraph) return paragraph;
  }
  return docxSchema.nodes.rawBlock.create({
    xml: serializeXml(el),
    name: el.nodeName,
  });
}

/**
 * The lines of one table, and the size of the grid they are laid over.
 * A cell reads its own four defaults out of these, according to where in the grid it sits.
 * The cell margins are the same for every cell, so they come along for the ride.
 */
interface TableLines extends GridSize {
  outer: TableFormat | null;
  inside: InsideBorders;
  margins: CellMargins;
}

/** The lines the four sides of one cell fall back on, merges included in where its edges lie */
function cellBorderDefaultsFor(
  draft: CellDraft,
  lines: TableLines
): CellBorderDefaults {
  return cellBorderDefaults(
    gridEdgesOf(
      {
        top: draft.row,
        bottom: draft.row + draft.rowspan,
        left: draft.col,
        right: draft.col + draft.colspan,
      },
      lines
    ),
    lines.outer,
    lines.inside
  );
}

/** Builds a cell. null if there is no block inside it at all */
function buildCell(
  draft: CellDraft,
  lines: TableLines,
  sources: ImportSources
): PMNode | null {
  const tcPr = childByLocalName(draft.el, "tcPr");
  const blocks: PMNode[] = [];
  for (const child of elementChildren(draft.el)) {
    if (child.localName === "tcPr") continue;
    blocks.push(buildCellBlock(child, sources));
  }
  if (blocks.length === 0) return null;

  return docxSchema.nodes.tableCell.create(
    {
      colspan: draft.colspan,
      rowspan: draft.rowspan,
      colwidth: null,
      tcAttrs: attrString(draft.el),
      tcPr: tcPr ? serializeXml(tcPr) : null,
      tcW: readTableWidth(tcPr, "tcW"),
      format: readCellFormat(
        tcPr,
        cellBorderDefaultsFor(draft, lines),
        lines.margins
      ),
      sdtPrefix: draft.control?.prefix ?? null,
      sdtContentsLocked: draft.control?.contentsLocked ?? false,
      sdtDeletionLocked: draft.control?.deletionLocked ?? false,
    },
    blocks
  );
}

function buildRow(
  row: RawRow,
  drafts: CellDraft[],
  lines: TableLines,
  sources: ImportSources
): PMNode | null {
  const cells: PMNode[] = [];
  for (const draft of drafts) {
    const cell = buildCell(draft, lines, sources);
    if (!cell) return null;
    cells.push(cell);
  }
  const trPr = childByLocalName(row.el, "trPr");
  return docxSchema.nodes.tableRow.create(
    {
      trAttrs: attrString(row.el),
      tblPrEx: row.tblPrEx ? serializeXml(row.tblPrEx) : null,
      trPr: trPr ? serializeXml(trPr) : null,
      format: readRowFormat(trPr),
    },
    cells
  );
}

/**
 * The style whose values lie underneath this table's own formatting.
 *
 * A table that points at no style, or at one that is not defined, falls back on the document's
 * default table style, which is what OOXML applies to an object with no style of its own.
 * A style that does resolve needs no fallback underneath it, because a real table style is based on
 * the default one anyway.
 *
 * Only what the style wrote in its `w:tblPr` reaches the table. Its conditional formatting
 * (`w:tblStylePr`, which dresses the header row, the first column and the banded rows differently),
 * the `w:tblLook` saying which of those parts are switched on, and the paragraph and run formatting
 * it carries for the text inside the cells are all left unread. Nothing is lost on the way out,
 * because export never rebuilds a style.
 */
function tableStyleOf(
  tblPr: Element | null,
  styles: StyleTable,
  defaultTableStyleId: string | null
): StyleFormat | undefined {
  const styleId = tblStyleIdOf(tblPr);
  return (
    (styleId !== null ? styles.get(styleId) : undefined) ??
    (defaultTableStyleId !== null ? styles.get(defaultTableStyleId) : undefined)
  );
}

/** Moves a `<w:tbl>` into a table node. null if it cannot be modelled */
export function buildTable(
  el: Element,
  srcId: number | null,
  sources: ImportSources = NO_IMPORT_SOURCES,
  styles: StyleTable = NO_STYLES,
  defaultTableStyleId: string | null = null
): PMNode | null {
  const parts = readTableParts(el);
  if (!parts) return null;

  const gridCols = readGridCols(parts.tblGrid);
  const width = gridWidthOf(parts.rows);
  if (width === null || width === 0) return null;
  // If there is a tblGrid, the grid and the cells have to add up to the same count
  if (gridCols.length > 0 && gridCols.length !== width) return null;

  const drafts = resolveVerticalMerges(parts.rows, width);
  if (!drafts) return null;

  const style = tableStyleOf(parts.tblPr, styles, defaultTableStyleId);
  const tableFormat = layerTableFormat(
    style?.table ?? {},
    readTableFormat(parts.tblPr)
  );
  const lines: TableLines = {
    rows: parts.rows.length,
    cols: width,
    outer: tableFormat,
    inside: layerInsideBorders(
      style?.tableInside ?? NO_INSIDE_BORDERS,
      readInsideBorders(parts.tblPr)
    ),
    margins: layerCellMargins(
      style?.tableCellMargins ?? NO_CELL_MARGINS,
      readCellMarginsOf(parts.tblPr, "tblCellMar")
    ),
  };

  const rows: PMNode[] = [];
  for (const [index, row] of parts.rows.entries()) {
    const built = buildRow(row, drafts[index], lines, sources);
    if (!built) return null;
    rows.push(built);
  }

  return docxSchema.nodes.table.create(
    {
      srcId,
      tblAttrs: attrString(el),
      tblPr: parts.tblPr ? serializeXml(parts.tblPr) : null,
      tblW: readTableWidth(parts.tblPr, "tblW"),
      gridCols,
      format: tableFormat,
      // The cells need these again whenever an edit derives their display values afresh
      styleInside: toInsideBorders(style?.tableInside),
      styleCellMargins: toCellMargins(style?.tableCellMargins),
    },
    rows
  );
}
