/**
 * The lines a cell draws because of where it sits in the grid, kept true after a structure edit.
 *
 * Every line of a table is drawn by its cells (see `docx/tableFormatting`), so which line a side
 * falls back on depends on where in the grid the cell sits. A new cell inherits its neighbour's
 * formatting, which is what a background has to do but not a line: the row appended under the last
 * one would draw the table's outer line against the row above it, and deleting the last row would
 * leave the table with no line along its bottom.
 *
 * So the cells of a table whose grid moved derive their display values again, along the same path
 * the import takes. What a cell wrote down itself lives in its `w:tcPr` and is read straight back
 * out of it, so only the share that came from the table can change.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";
import { Transform } from "prosemirror-transform";
import {
  type CellBorderDefaults,
  cellBorderDefaults,
  cellMarginsOf,
  gridEdgesOf,
  insideBordersOf,
  layerCellMargins,
  layerInsideBorders,
  NO_BORDER_DEFAULTS,
  NO_CELL_MARGINS,
  NO_INSIDE_BORDERS,
  readCellProps,
} from "../docx/tableFormatting";
import {
  type CellFormat,
  type CellMargins,
  type InsideBorders,
  spanCount,
  type TableFormat,
  toCellFormat,
  toCellMargins,
  toInsideBorders,
  toTableFormat,
} from "../model/format";
import type { NodeAttrs, TableGridMap } from "./format";

function text(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

/**
 * What a table lays down for its cells to draw.
 * The values a table style laid down are not in the `tblPr`, so the table carries them separately.
 */
export interface TableCellSources {
  outer: TableFormat | null;
  inside: InsideBorders;
  margins: CellMargins;
}

export function tableCellSources(table: PMNode): TableCellSources {
  const tblPr = text(table.attrs.tblPr);
  return {
    outer: toTableFormat(table.attrs.format),
    inside: layerInsideBorders(
      toInsideBorders(table.attrs.styleInside) ?? NO_INSIDE_BORDERS,
      insideBordersOf(tblPr)
    ),
    margins: layerCellMargins(
      toCellMargins(table.attrs.styleCellMargins) ?? NO_CELL_MARGINS,
      cellMarginsOf(tblPr)
    ),
  };
}

/** The lines the cell at this position falls back on for the sides it draws no border of its own on */
export function cellDefaultsAt(
  map: TableGridMap,
  pos: number,
  sources: TableCellSources
): CellBorderDefaults {
  return cellBorderDefaults(
    gridEdgesOf(map.findCell(pos), { rows: map.height, cols: map.width }),
    sources.outer,
    sources.inside
  );
}

const CELL_FORMAT_KEYS = [
  "borderTop",
  "borderBottom",
  "borderLeft",
  "borderRight",
  "background",
  "verticalAlign",
  "paddingTopPt",
  "paddingRightPt",
  "paddingBottomPt",
  "paddingLeftPt",
] as const;

type BorderFormatKey =
  | "borderTop"
  | "borderBottom"
  | "borderLeft"
  | "borderRight";

function sameCellFormat(a: CellFormat | null, b: CellFormat | null): boolean {
  if (!a || !b) return a === b;
  return CELL_FORMAT_KEYS.every((key) => a[key] === b[key]);
}

interface CellFix {
  /** The position of the cell within the table's content */
  pos: number;
  attrs: NodeAttrs;
}

interface DerivedCell {
  pos: number;
  attrs: NodeAttrs;
  /** Mutable display values; they never replace the cell's preserved OOXML. */
  format: CellFormat;
  /** Only properties written directly in this cell's `w:tcPr`. */
  direct: CellFormat | null;
}

function hasDirectBorder(cell: DerivedCell, key: BorderFormatKey): boolean {
  return cell.direct?.[key] !== undefined;
}

/**
 * Suppresses the inherited side of a shared edge when its opposite is a direct cell border.
 * ECMA-376 Part 1 §17.4.66 gives the direct border precedence; leaving it as the only CSS candidate
 * also keeps one segment from repainting the whole side of an adjacent merged cell.
 */
function reconcileSharedBorder(
  first: DerivedCell,
  firstKey: BorderFormatKey,
  second: DerivedCell,
  secondKey: BorderFormatKey
): void {
  const firstIsDirect = hasDirectBorder(first, firstKey);
  const secondIsDirect = hasDirectBorder(second, secondKey);
  if (firstIsDirect === secondIsDirect) return;
  if (firstIsDirect) {
    second.format[secondKey] = "none";
  } else {
    first.format[firstKey] = "none";
  }
}

function reconcileSharedBorders(
  map: TableGridMap,
  cells: ReadonlyMap<number, DerivedCell>
): void {
  for (let row = 0; row < map.height; row += 1) {
    for (let col = 1; col < map.width; col += 1) {
      const leftPos = map.map[row * map.width + col - 1];
      const rightPos = map.map[row * map.width + col];
      if (leftPos === rightPos) continue;
      const left = cells.get(leftPos);
      const right = cells.get(rightPos);
      if (left && right) {
        reconcileSharedBorder(left, "borderRight", right, "borderLeft");
      }
    }
  }
  for (let row = 1; row < map.height; row += 1) {
    for (let col = 0; col < map.width; col += 1) {
      const topPos = map.map[(row - 1) * map.width + col];
      const bottomPos = map.map[row * map.width + col];
      if (topPos === bottomPos) continue;
      const top = cells.get(topPos);
      const bottom = cells.get(bottomPos);
      if (top && bottom) {
        reconcileSharedBorder(top, "borderBottom", bottom, "borderTop");
      }
    }
  }
}

/** The cells of this table whose display values do not match the spot they now sit in */
function cellFixes(table: PMNode): CellFix[] {
  const map = TableMap.get(table);
  const sources = tableCellSources(table);
  const cells = new Map<number, DerivedCell>();
  // A merged cell is pointed at from every spot it covers, so each cell is looked at once
  for (const pos of new Set(map.map)) {
    const cell = table.nodeAt(pos);
    if (!cell) continue;
    const tcPr = text(cell.attrs.tcPr);
    const format = readCellProps(
      tcPr,
      cellDefaultsAt(map, pos, sources),
      sources.margins
    );
    cells.set(pos, {
      pos,
      attrs: cell.attrs,
      format: { ...format },
      direct: readCellProps(tcPr, NO_BORDER_DEFAULTS),
    });
  }
  reconcileSharedBorders(map, cells);

  const fixes: CellFix[] = [];
  for (const cell of cells.values()) {
    const format = Object.keys(cell.format).length === 0 ? null : cell.format;
    if (sameCellFormat(format, toCellFormat(cell.attrs.format))) continue;
    fixes.push({ pos: cell.pos, attrs: { ...cell.attrs, format } });
  }
  return fixes;
}

interface PlacedTable {
  /** The position of the table itself */
  pos: number;
  table: PMNode;
}

/**
 * Every table of the document, in the order they stand.
 * The text inside a paragraph is not walked into: this runs after every edit that changed the
 * document, typing included, so it may not cost more than a walk over the blocks.
 */
function tablesOf(doc: PMNode): PlacedTable[] {
  const tables: PlacedTable[] = [];
  doc.descendants((node, pos) => {
    if (node.type.spec.tableRole === "table") tables.push({ pos, table: node });
    return !node.isTextblock;
  });
  return tables;
}

function sameSpans(a: PMNode, b: PMNode): boolean {
  return (
    spanCount(a.attrs.colspan) === spanCount(b.attrs.colspan) &&
    spanCount(a.attrs.rowspan) === spanCount(b.attrs.rowspan)
  );
}

function sameRow(a: PMNode, b: PMNode): boolean {
  return (
    a.childCount === b.childCount &&
    a.children.every((cell, index) => sameSpans(cell, b.child(index)))
  );
}

/** The grid is all a cell's lines depend on, so a table whose text alone was edited is left alone */
function sameGrid(a: PMNode, b: PMNode): boolean {
  return (
    a.childCount === b.childCount &&
    a.children.every((row, index) => sameRow(row, b.child(index)))
  );
}

function sameCellFormattingInputs(a: PMNode, b: PMNode): boolean {
  return a.children.every((row, rowIndex) =>
    row.children.every(
      (cell, cellIndex) =>
        cell.attrs.tcPr === b.child(rowIndex).child(cellIndex).attrs.tcPr
    )
  );
}

function sameFormattingInputs(a: PMNode, b: PMNode): boolean {
  return (
    sameGrid(a, b) &&
    a.attrs.tblPr === b.attrs.tblPr &&
    a.attrs.format === b.attrs.format &&
    a.attrs.styleInside === b.attrs.styleInside &&
    a.attrs.styleCellMargins === b.attrs.styleCellMargins &&
    sameCellFormattingInputs(a, b)
  );
}

/** Derives shared-border display values before the first editor view is drawn. */
export function withDerivedGridBorders(doc: PMNode): PMNode {
  const transform = new Transform(doc);
  for (const { pos, table } of tablesOf(doc)) {
    for (const fix of cellFixes(table)) {
      transform.setNodeMarkup(pos + 1 + fix.pos, null, fix.attrs);
    }
  }
  return transform.doc;
}

/**
 * Keeps the derived lines current after a table's grid or border inputs change.
 *
 * The tables are paired up by the order they stand in. A document that gained or lost a whole table
 * pairs the ones after it with the wrong partner, which costs a derivation that changes no cell.
 *
 * No transaction of its own goes to the history: ProseMirror hands an appended transaction to the
 * history as part of the same event, so a single undo takes the edit and this correction back together.
 */
export function gridBorders(): Plugin {
  return new Plugin({
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) {
        return null;
      }
      const before = tablesOf(oldState.doc);
      const tr = newState.tr;
      tablesOf(newState.doc).forEach(({ pos, table }, index) => {
        const was = before[index]?.table;
        if (was && sameFormattingInputs(was, table)) return;
        for (const fix of cellFixes(table)) {
          tr.setNodeMarkup(pos + 1 + fix.pos, null, fix.attrs);
        }
      });
      return tr.docChanged ? tr : null;
    },
  });
}
