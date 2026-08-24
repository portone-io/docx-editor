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
import { docxSchema } from "../schema";
import { A4_PORTRAIT, bodyWidth, type PageGeometry } from "./pageGeometry";
import { parsePropsXml } from "./propsXml";
import {
  type CellBorderDefaults,
  cellBorderDefaults,
  gridEdgesOf,
  NO_CELL_MARGINS,
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

/** The four outer sides and the two lines between cells. The order is the one CT_TblBorders lays down */
const BORDER_SIDES = [
  "top",
  "left",
  "bottom",
  "right",
  "insideH",
  "insideV",
] as const;

function borderXml(side: string): string {
  return `<w:${side} w:val="single" w:sz="${BORDER_EIGHTHS}" w:space="0" w:color="000000"/>`;
}

/**
 * The order of the children is the one CT_TblPr lays down.
 *
 * The widths are pinned to the grid (`tblLayout`).
 * Our screen always draws to the grid, so writing it this way makes Word draw the same widths.
 */
function tablePropsXml(width: number): string {
  const borders = BORDER_SIDES.map(borderXml).join("");
  return (
    "<w:tblPr>" +
    `<w:tblW w:w="${width}" w:type="dxa"/>` +
    `<w:tblBorders>${borders}</w:tblBorders>` +
    '<w:tblLayout w:type="fixed"/>' +
    "</w:tblPr>"
  );
}

function cellPropsXml(width: number): string {
  return `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>`;
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
function cellTemplate(
  width: number,
  defaults: CellBorderDefaults
): CellTemplate {
  const tcPr = cellPropsXml(width);
  const el = parsePropsXml(tcPr);
  return {
    tcPr,
    tcW: readTableWidth(el, "tcW"),
    // A new table writes no `tblCellMar`, so its cells take their padding from the stylesheet
    format: readCellFormat(el, defaults, NO_CELL_MARGINS),
  };
}

function cellNode(props: CellTemplate): PMNode {
  return docxSchema.nodes.tableCell.create(
    {
      colspan: 1,
      rowspan: 1,
      colwidth: null,
      tcAttrs: null,
      tcPr: props.tcPr,
      tcW: props.tcW,
      format: props.format,
    },
    docxSchema.nodes.paragraph.create()
  );
}

function rowNode(cells: readonly CellTemplate[]): PMNode {
  return docxSchema.nodes.tableRow.create(
    { trAttrs: null, trPr: null, format: null },
    cells.map(cellNode)
  );
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
  const total = bodyWidth(geometry).twips;
  const gridCols = evenGridCols(cols, total);
  const tblPr = tablePropsXml(total);
  const tblPrEl = parsePropsXml(tblPr);
  const inside = readInsideBorders(tblPrEl);
  const outer = readTableFormat(tblPrEl);

  const cellAt = (row: number, col: number) =>
    cellTemplate(
      gridCols[col],
      cellBorderDefaults(
        gridEdgesOf(
          { top: row, bottom: row + 1, left: col, right: col + 1 },
          { rows, cols }
        ),
        outer,
        inside
      )
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
    Array.from({ length: rows }, (_, row) =>
      rowNode(gridCols.map((_width, col) => cellAt(row, col)))
    )
  );
}
