/**
 * Imports one OOXML table, converting vertical merge continuations to rowspans and layering style
 * values for display. Unsupported grids return null so the caller can preserve their source XML.
 */

import type { Node as PMNode } from "prosemirror-model";
import { childValue, wAttr } from "../ooxml/units";
import {
  attrString,
  childByLocalName,
  elementChildren,
  serializeXml,
} from "../ooxml/xml";
import { docxSchema } from "../schema";
import {
  type ControlFacts,
  controlAttrs,
  NO_CONTROL,
  WRAPPED_CONTROL_ATTRS,
} from "../schema/controlAttrs";
import {
  type FormattingContext,
  layerTableFormat,
  NO_FORMATTING,
  type ParagraphPlacement,
  styledParagraph,
  tableStyleAttrs,
  tableStyleFor,
} from "./formatting";
import {
  buildParagraph,
  type ImportSources,
  NO_IMPORT_SOURCES,
} from "./importParagraph";
import { policyFor } from "./importPolicy";
import { buildPreservedBlock } from "./importPreserved";
import { buildSdtBlock } from "./importSdtBlock";
import { controlFactsFrom, readSdtWrapper } from "./sdt";
import {
  cellConditionsOf,
  cellDefaultsFor,
  type GridRect,
  type GridSize,
  layerBandSizes,
  layerCellMargins,
  layerInsideBorders,
  NO_BAND_SIZES,
  NO_CELL_MARGINS,
  NO_INSIDE_BORDERS,
  readBandSizes,
  readCellFormat,
  readCellMarginsOf,
  readGridCols,
  readInsideBorders,
  readRowFormat,
  readTableFormat,
  readTableLook,
  readTableWidth,
  type TableCellSources,
  tblStyleIdOf,
} from "./tableFormatting";

type VerticalMerge = "restart" | "continue" | null;

/**
 * The content control around one cell or one row, as that node carries it on
 * (`schema/controlAttrs`). `CT_SdtCell` and `CT_SdtRow` say the same things about themselves, so
 * the two are read and carried alike.
 */
type WrappedControl = ControlFacts;

interface RawCell {
  el: Element;
  gridSpan: number;
  vMerge: VerticalMerge;
  control: WrappedControl | null;
  /** The markers that stood between this cell and the next one */
  trailingXml: string | null;
}

interface RawRow {
  el: Element;
  tblPrEx: Element | null;
  cells: RawCell[];
  /** The content control the file put around the whole row, or null where none did */
  control: WrappedControl | null;
  /** The markers that stood ahead of the row's first cell */
  leadingXml: string | null;
  /** The markers that stood between this row and the next one */
  trailingXml: string | null;
}

/** One cell to be built. Its rowspan grows as the continuing cells are counted */
interface CellDraft {
  el: Element;
  /** The row and the column the cell starts at, which is what tells it where its lines come from */
  row: number;
  col: number;
  colspan: number;
  rowspan: number;
  control: WrappedControl | null;
  trailingXml: string | null;
}

/**
 * Whether an element at this level is carried along rather than read or demoted over.
 *
 * `w:tbl` and `w:tr` are the two levels with no node to keep a stranger in (`./importPolicy`), so
 * what a rule there says is invisible rides on the child before it and goes back out from there.
 */
function isTransparent(el: Element, level: "tbl" | "tr"): boolean {
  const tier = policyFor(el, level).tier;
  return tier === "marker" || tier === "ignorable";
}

/**
 * Everything gathered since the last time it was emptied, as one string, and nothing next time.
 *
 * A run of them is one value because they stand together: the end of one bookmark and the start
 * of the next have nothing between them.
 */
function takeGathered(gathered: string[]): string | null {
  return gathered.length === 0 ? null : gathered.splice(0).join("");
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

/** A cell or a row taken out of the content control that wrapped it, plus that control */
interface WrappedNode {
  el: Element;
  control: WrappedControl;
}

/**
 * Takes the single `w:tc` or `w:tr` out of a `w:sdt` content control (`CT_SdtCell` §17.5.2.32,
 * `CT_SdtRow` §17.5.2.30).
 *
 * On top of the wrapper shape `readSdtWrapper` insists on, the content has to hold nothing but
 * one element of that name - not two of them, and not another control. Everything else is null,
 * which leaves the whole table preserved (`spec/notes/contentControls.md`).
 */
function readSdtSingle(el: Element, localName: string): WrappedNode | null {
  const wrapper = readSdtWrapper(el);
  if (!wrapper) return null;

  const inner = elementChildren(wrapper.content);
  if (inner.length !== 1 || inner[0].localName !== localName) return null;

  return { el: inner[0], control: controlFactsFrom(wrapper) };
}

/**
 * A cell that starts a vertical merge is a cell of the model in its own right, and `serializeTable`
 * puts its wrapper back around it. The cells that only continue the merge are created fresh on
 * export instead, and a wrapper of their own could not be recreated for them.
 */
function readSdtCell(el: Element): WrappedNode | null {
  const cell = readSdtSingle(el, "tc");
  if (!cell) return null;
  return readVerticalMerge(childByLocalName(cell.el, "tcPr")) === "continue"
    ? null
    : cell;
}

/**
 * A row all of whose cells only continue a vertical merge is no row of the model at all
 * (`resolveVerticalMerges`), so a control around one stands the table down there instead of here.
 */
function readSdtRow(el: Element): WrappedNode | null {
  return readSdtSingle(el, "tr");
}

/**
 * Hangs the markers gathered so far off the cell they stood after. false when there is no cell to
 * hang them on.
 *
 * The cells that only continue a vertical merge are made fresh on export rather than written from
 * a node of the model, so a marker standing after one has nowhere to go back to and the table is
 * stood down instead of quietly losing it.
 */
function trailCell(cells: RawCell[], gathered: string[]): boolean {
  const xml = takeGathered(gathered);
  if (xml === null) return true;
  const last = cells.at(-1);
  if (!last || last.vMerge === "continue") return false;
  last.trailingXml = (last.trailingXml ?? "") + xml;
  return true;
}

/**
 * Reads a single row. null if there is any child other than `w:tc`, `w:trPr`, `w:tblPrEx` and
 * what `./importPolicy` carries along at this level, or a content control we could not write out.
 *
 * A `w:tblPrEx` states the table properties this one row departs from. We do not read it, so
 * the row is drawn with the table's own values, but it is carried along to go back out untouched.
 */
function readRow(el: Element, control: WrappedControl | null): RawRow | null {
  const cells: RawCell[] = [];
  const gathered: string[] = [];
  let tblPrEx: Element | null = null;
  let leadingXml: string | null = null;
  for (const child of elementChildren(el)) {
    if (child.localName === "trPr") continue;
    if (child.localName === "tblPrEx") {
      tblPrEx = child;
      continue;
    }
    if (isTransparent(child, "tr")) {
      gathered.push(serializeXml(child));
      continue;
    }
    const sdt = child.localName === "sdt" ? readSdtCell(child) : null;
    if (!sdt && child.localName !== "tc") return null;
    // What stood ahead of the first cell belongs to the row; the rest to the cell before it
    if (cells.length === 0) leadingXml = takeGathered(gathered);
    else if (!trailCell(cells, gathered)) return null;
    const tc = sdt ? sdt.el : child;
    const tcPr = childByLocalName(tc, "tcPr");
    cells.push({
      el: tc,
      gridSpan: readGridSpan(tcPr),
      vMerge: readVerticalMerge(tcPr),
      control: sdt ? sdt.control : null,
      trailingXml: null,
    });
  }
  if (cells.length === 0 || !trailCell(cells, gathered)) return null;
  return { el, tblPrEx, cells, control, leadingXml, trailingXml: null };
}

interface TableParts {
  tblPr: Element | null;
  tblGrid: Element | null;
  /** The `w:tblGridChange` the grid closed with, as it stood */
  gridChange: string | null;
  rows: RawRow[];
  /** The markers that stood ahead of the first row */
  leadingXml: string | null;
}

/** Keeps the revision's namespace context when its parent grid is rebuilt. */
function gridRevisionXml(grid: Element): string | null {
  const revision = childByLocalName(grid, "tblGridChange");
  if (revision === null) return null;
  // The rebuilt grid no longer supplies these bindings. Keep all of them, including prefixes
  // used only in QName-valued attributes, without overriding declarations inside the revision.
  const declarations = Array.from(grid.attributes).filter(
    (attr) =>
      attr.namespaceURI === "http://www.w3.org/2000/xmlns/" &&
      !revision.hasAttribute(attr.name)
  );
  if (declarations.length === 0) return serializeXml(revision);

  const preserved = revision.ownerDocument.importNode(revision, true);
  for (const attr of declarations) {
    preserved.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
  }
  return serializeXml(preserved);
}

/**
 * Splits a table into the pieces it is made of. null if a child we do not know is mixed in.
 *
 * A marker standing between two rows rides on the row before it, and one ahead of the first row on
 * the table itself, so a bookmark spanning a column no longer stands the whole table down.
 *
 * A `w:sdt` here wraps one row (`CT_SdtRow`), which `readSdtRow` reads onto that row.
 */
function readTableParts(el: Element): TableParts | null {
  let tblPr: Element | null = null;
  let tblGrid: Element | null = null;
  let gridChange: string | null = null;
  const rows: RawRow[] = [];
  const gathered: string[] = [];
  let leadingXml: string | null = null;
  for (const child of elementChildren(el)) {
    if (child.localName === "tblPr") {
      tblPr = child;
      continue;
    }
    if (child.localName === "tblGrid") {
      tblGrid = child;
      gridChange = gridRevisionXml(child);
      continue;
    }
    if (isTransparent(child, "tbl")) {
      gathered.push(serializeXml(child));
      continue;
    }
    const sdt = child.localName === "sdt" ? readSdtRow(child) : null;
    if (!sdt && child.localName !== "tr") return null;
    const row = readRow(sdt?.el ?? child, sdt?.control ?? null);
    if (!row) return null;
    const previous = rows.at(-1);
    if (previous) previous.trailingXml = takeGathered(gathered);
    else leadingXml = takeGathered(gathered);
    rows.push(row);
  }
  const last = rows.at(-1);
  if (!last) return null;
  last.trailingXml = takeGathered(gathered);
  return { tblPr, tblGrid, gridChange, rows, leadingXml };
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
          trailingXml: cell.trailingXml,
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

/**
 * A single block inside a cell, dressed by the parts of the table the cell belongs to.
 *
 * A control opens as the container of the blocks it holds, whose paragraphs are dressed the same
 * way the cell's own are. A block that is neither is kept as the XML it came as, drawn by whatever
 * `./importPolicy` says is on screen of it at the level it stands at.
 */
function buildCellBlock(
  el: Element,
  sources: ImportSources,
  context: FormattingContext,
  placement: ParagraphPlacement,
  level: "tc" | "sdtContent" = "tc"
): PMNode {
  if (el.localName === "p") {
    return styledParagraph(
      buildParagraph(el, null, sources),
      context,
      placement
    );
  }
  if (el.localName === "sdt") {
    const control = buildSdtBlock(el, null, (child) =>
      buildCellBlock(child, sources, context, placement, "sdtContent")
    );
    if (control) return control;
  }
  return buildPreservedBlock(el, null, level);
}

/** The block of the grid one cell covers, merges included */
function cellRect(draft: CellDraft): GridRect {
  return {
    top: draft.row,
    bottom: draft.row + draft.rowspan,
    left: draft.col,
    right: draft.col + draft.colspan,
  };
}

/** Builds a cell. null if there is no block inside it at all */
function buildCell(
  draft: CellDraft,
  table: TableCells,
  sources: ImportSources,
  context: FormattingContext
): PMNode | null {
  const tcPr = childByLocalName(draft.el, "tcPr");
  const defaults = cellDefaultsFor(cellRect(draft), table.grid, table.sources);
  const blocks: PMNode[] = [];
  for (const child of elementChildren(draft.el)) {
    if (child.localName === "tcPr") continue;
    blocks.push(buildCellBlock(child, sources, context, defaults.placement));
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
      format: readCellFormat(tcPr, defaults),
      ...controlAttrs(WRAPPED_CONTROL_ATTRS, draft.control ?? NO_CONTROL),
      trailingXml: draft.trailingXml,
    },
    blocks
  );
}

/** The size of a table's grid and what it lays down for the cells laid over it */
interface TableCells {
  grid: GridSize;
  sources: TableCellSources;
}

function buildRow(
  row: RawRow,
  drafts: CellDraft[],
  table: TableCells,
  sources: ImportSources,
  context: FormattingContext
): PMNode | null {
  const cells: PMNode[] = [];
  for (const draft of drafts) {
    const cell = buildCell(draft, table, sources, context);
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
      ...controlAttrs(WRAPPED_CONTROL_ATTRS, row.control ?? NO_CONTROL),
      leadingXml: row.leadingXml,
      trailingXml: row.trailingXml,
    },
    cells
  );
}

/** Moves a `<w:tbl>` into a table node. null if it cannot be modelled */
export function buildTable(
  el: Element,
  srcId: string | null,
  sources: ImportSources = NO_IMPORT_SOURCES,
  context: FormattingContext = NO_FORMATTING
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

  const styleId = tblStyleIdOf(parts.tblPr);
  const style = tableStyleFor(styleId, context);
  const tableFormat = layerTableFormat(
    style?.table ?? {},
    readTableFormat(parts.tblPr)
  );
  // A band size is normally written in the style, but a table may state one of its own
  const bands = layerBandSizes(
    style?.tableBands ?? NO_BAND_SIZES,
    readBandSizes(parts.tblPr)
  );
  const conditions = cellConditionsOf(style?.tableConditions ?? {});
  const table: TableCells = {
    grid: { rows: parts.rows.length, cols: width },
    sources: {
      outer: tableFormat,
      inside: layerInsideBorders(
        style?.tableInside ?? NO_INSIDE_BORDERS,
        readInsideBorders(parts.tblPr)
      ),
      margins: layerCellMargins(
        style?.tableCellMargins ?? NO_CELL_MARGINS,
        readCellMarginsOf(parts.tblPr, "tblCellMar")
      ),
      look: readTableLook(parts.tblPr),
      bands,
      conditions,
      styleId,
    },
  };

  const rows: PMNode[] = [];
  for (const [index, row] of parts.rows.entries()) {
    const built = buildRow(row, drafts[index], table, sources, context);
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
      gridChange: parts.gridChange,
      leadingXml: parts.leadingXml,
      ...tableStyleAttrs(parts.tblPr, context),
    },
    rows
  );
}
