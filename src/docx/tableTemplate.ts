/**
 * Builds a single new table as a document model.
 *
 * The editor knows nothing about XML, so the `w:tblPr` and `w:tcPr` the new table will use are
 * written out as text here too. The fragments we write are then read back with the same reading
 * functions used when opening a document, to derive the display values.
 * That way a freshly inserted table looks on screen exactly as it does after saving and reopening.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { CellFormat, TableWidth } from "../model/format";
import { elementXml } from "../ooxml/element";
import { wName } from "../ooxml/names";
import { orderedElement, parsePropsXml } from "../ooxml/props";
import { docxSchema } from "../schema";
import { A4_PORTRAIT, bodyWidth, type PageGeometry } from "./pageGeometry";
import {
  type CellDefaults,
  cellDefaultsFor,
  type GridRect,
  NO_CELL_SOURCES,
  readCellFormat,
  readInsideBorders,
  readTableFormat,
  readTableWidth,
} from "./tableFormatting";

/**
 * The limit on how many rows and cells one table can hold.
 * The toolbar grid lets you pick from something smaller (6x6); this only blocks absurd values.
 */
const MAX_TABLE_SIDE = 50;

/** Whether the value can be used as a row count or a cell count */
export function isTableSide(value: number): boolean {
  return Number.isInteger(value) && value >= 1 && value <= MAX_TABLE_SIDE;
}

/** The border thickness of a new table. `w:sz` is in 1/8 of a point, so 4 is 0.5pt */
const BORDER_EIGHTHS = 4;

/** The four outer sides and the two lines between cells */
const BORDER_SIDES = [
  "top",
  "left",
  "bottom",
  "right",
  "insideH",
  "insideV",
] as const;

function borderXml(side: string): string {
  return elementXml(wName(side), [
    [wName("val"), "single"],
    [wName("sz"), `${BORDER_EIGHTHS}`],
    [wName("space"), "0"],
    [wName("color"), "000000"],
  ]);
}

/**
 * The widths are pinned to the grid (`tblLayout`).
 * Our screen always draws to the grid, so writing it this way makes Word draw the same widths.
 */
function tablePropsXml(width: number): string {
  const borders = orderedElement(
    wName("tblBorders"),
    [],
    BORDER_SIDES.map((side) => ({ name: side, xml: borderXml(side) }))
  );
  return orderedElement(
    wName("tblPr"),
    [],
    [
      { name: "tblW", xml: widthXml("tblW", width) },
      { name: "tblBorders", xml: borders },
      {
        name: "tblLayout",
        xml: elementXml(wName("tblLayout"), [[wName("type"), "fixed"]]),
      },
    ]
  );
}

/** A width pinned to the grid, in twips */
function widthXml(name: "tblW" | "tcW", width: number): string {
  return elementXml(wName(name), [
    [wName("w"), `${width}`],
    [wName("type"), "dxa"],
  ]);
}

function cellPropsXml(width: number): string {
  return orderedElement(
    wName("tcPr"),
    [],
    [{ name: "tcW", xml: widthXml("tcW", width) }]
  );
}

/**
 * Column widths that divide the body width evenly.
 * The remainder is handed out one twip at a time starting from the first column, so the sum always
 * equals the body width.
 */
function evenGridCols(cols: number, total: number): number[] {
  const base = Math.floor(total / cols);
  const extra = total - base * cols;
  return Array.from({ length: cols }, (_, index) =>
    index < extra ? base + 1 : base
  );
}

/** The formatting fragment one cell uses, and the display values read out of that fragment */
interface CellTemplate {
  tcPr: string;
  tcW: TableWidth | null;
  format: CellFormat | null;
}

/**
 * One cell of the new table.
 * Which of the table's lines its four sides fall on depends on where in the grid it sits, so the
 * fragment is read back once per cell rather than once per column.
 */
function cellTemplate(width: number, defaults: CellDefaults): CellTemplate {
  const tcPr = cellPropsXml(width);
  const el = parsePropsXml(tcPr);
  return {
    tcPr,
    tcW: readTableWidth(el, "tcW"),
    // A new table writes no `tblCellMar`, so its cells take their padding from the stylesheet
    format: readCellFormat(el, defaults),
  };
}

function cellNode(props: CellTemplate, plan: TableCellPlan): PMNode {
  const content = plan.content ?? [];
  return docxSchema.nodes.tableCell.create(
    {
      colspan: plan.rect.right - plan.rect.left,
      rowspan: plan.rect.bottom - plan.rect.top,
      colwidth: null,
      tcAttrs: null,
      tcPr: props.tcPr,
      tcW: props.tcW,
      format: props.format,
    },
    // A cell holds blocks, so one holding nothing still holds an empty paragraph
    content.length === 0 ? docxSchema.nodes.paragraph.create() : [...content]
  );
}

function rowNode(cells: readonly PMNode[]): PMNode {
  return docxSchema.nodes.tableRow.create(
    { trAttrs: null, trPr: null, format: null },
    [...cells]
  );
}

/**
 * One cell of a table being built: the block of the grid it covers, and the blocks it holds.
 *
 * A cell reaching across columns or down rows exists once, as the cell it starts at, which is how
 * the model holds a merge and what `serializeTable` writes `w:gridSpan` and `w:vMerge` from.
 */
export interface TableCellPlan {
  rect: GridRect;
  content?: readonly PMNode[];
}

/** A table laid out by its caller: how big its grid is, and which cells start in each row */
export interface TablePlan {
  rows: number;
  cols: number;
  cells: readonly (readonly TableCellPlan[])[];
}

/** The plan of a table of single cells, which is what the toolbar inserts */
function evenPlan(rows: number, cols: number): TablePlan {
  return {
    rows,
    cols,
    cells: Array.from({ length: rows }, (_, row) =>
      Array.from({ length: cols }, (_, col) => ({
        rect: { top: row, bottom: row + 1, left: col, right: col + 1 },
      }))
    ),
  };
}

/**
 * A single table made up of nothing but empty cells.
 *
 * The table is as wide as one line of body text on the paper the document names, its column
 * widths divide that width evenly, and the borders are black 0.5pt solid lines. A document that
 * names no paper gets the A4 body width, which is what every document used to get.
 * Being a new table with no original fragment, it is rewritten in full by `serializeTable` on export.
 * The row and cell counts must be values that have passed through `isTableSide`.
 */
export function createTableNode(
  rows: number,
  cols: number,
  geometry: PageGeometry = A4_PORTRAIT
): PMNode {
  return createTableNodeFrom(evenPlan(rows, cols), geometry);
}

/**
 * The same table, laid out and filled by the caller: the shape a table read off the clipboard
 * takes, where the cells are neither all one wide nor all empty.
 *
 * Everything the table itself wears is the same as a newly inserted one's, so a pasted table is
 * drawn and written exactly like a table the toolbar put there. The grid size the plan states is
 * what the cells are measured and dressed against, so it must cover every cell the plan places.
 */
export function createTableNodeFrom(
  plan: TablePlan,
  geometry: PageGeometry = A4_PORTRAIT
): PMNode {
  const { rows, cols } = plan;
  const total = bodyWidth(geometry).twips;
  const gridCols = evenGridCols(cols, total);
  const tblPr = tablePropsXml(total);
  const tblPrEl = parsePropsXml(tblPr);
  const outer = readTableFormat(tblPrEl);
  // A new table wears no style, so nothing but its own lines reaches its cells
  const sources = {
    ...NO_CELL_SOURCES,
    outer,
    inside: readInsideBorders(tblPrEl),
  };

  const cellAt = (cell: TableCellPlan) =>
    cellNode(
      cellTemplate(
        gridCols
          .slice(cell.rect.left, cell.rect.right)
          .reduce((width, column) => width + column, 0),
        cellDefaultsFor(cell.rect, { rows, cols }, sources)
      ),
      cell
    );

  return docxSchema.nodes.table.create(
    {
      srcId: null,
      tblAttrs: null,
      tblPr,
      tblW: readTableWidth(tblPrEl, "tblW"),
      gridCols,
      format: outer,
    },
    plan.cells.map((row) => rowNode(row.map(cellAt)))
  );
}
