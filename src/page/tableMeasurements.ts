/** Measures safe row boundaries without treating pagination decorations as document content. */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { toRowFormat } from "../model/format";
import { editorAttributes } from "../styles/classNames";
import type { MeasuredTable, TableBoundary } from "./pageLayout";

interface RowEntry {
  node: PMNode;
  pos: number;
  top: number;
  height: number;
}

export interface TableMeasure {
  table: MeasuredTable;
  /** Height already added by the previous pagination pass */
  appliedHeight: number;
}

function span(value: unknown): number {
  return typeof value === "number" && Number.isInteger(value) && value > 0
    ? value
    : 1;
}

function paginationRows(table: HTMLElement): HTMLElement[] {
  return Array.from(
    table.querySelectorAll<HTMLElement>(
      `[${editorAttributes.tablePageSpace}],[${editorAttributes.tableRepeatedHeader}]`
    )
  );
}

function before(candidate: HTMLElement, row: HTMLElement): boolean {
  return Boolean(
    candidate.compareDocumentPosition(row) & Node.DOCUMENT_POSITION_FOLLOWING
  );
}

function rowEntries(
  view: EditorView,
  tableNode: PMNode,
  tablePos: number,
  tableDom: HTMLElement,
  appliedRows: readonly HTMLElement[],
  scale: number
): RowEntry[] {
  const entries: RowEntry[] = [];
  const tableTop = tableDom.getBoundingClientRect().top;
  tableNode.forEach((row, offset) => {
    const pos = tablePos + 1 + offset;
    const dom = view.nodeDOM(pos);
    if (!(dom instanceof HTMLElement)) return;
    const rect = dom.getBoundingClientRect();
    const appliedBefore = appliedRows.reduce(
      (total, candidate) =>
        total +
        (before(candidate, dom)
          ? candidate.getBoundingClientRect().height / scale
          : 0),
      0
    );
    entries.push({
      node: row,
      pos,
      top: (rect.top - tableTop) / scale - appliedBefore,
      height: rect.height / scale,
    });
  });
  return entries;
}

function unsafeBoundaries(rows: readonly RowEntry[]): ReadonlySet<number> {
  const unsafe = new Set<number>();
  rows.forEach((row, rowIndex) => {
    row.node.forEach((cell) => {
      const bottom = Math.min(rows.length, rowIndex + span(cell.attrs.rowspan));
      for (let boundary = rowIndex + 1; boundary < bottom; boundary += 1) {
        unsafe.add(boundary);
      }
    });
  });
  return unsafe;
}

function headerCount(rows: readonly RowEntry[]): number {
  let count = 0;
  while (
    count < rows.length &&
    toRowFormat(rows[count]?.node.attrs.format)?.repeatHeader === true
  ) {
    count += 1;
  }
  return count;
}

function columnCount(tableNode: PMNode): number {
  const grid = tableNode.attrs.gridCols;
  if (Array.isArray(grid) && grid.length > 0) return grid.length;
  const first = tableNode.firstChild;
  if (!first) return 1;
  let columns = 0;
  first.forEach((cell) => {
    columns += span(cell.attrs.colspan);
  });
  return Math.max(1, columns);
}

export function measureTable(
  view: EditorView,
  tableNode: PMNode,
  tablePos: number,
  tableDom: HTMLElement,
  scale = 1
): TableMeasure | null {
  if (tableNode.type.spec.tableRole !== "table") return null;
  const appliedRows = paginationRows(tableDom);
  const rows = rowEntries(
    view,
    tableNode,
    tablePos,
    tableDom,
    appliedRows,
    scale
  );
  if (rows.length === 0) return null;

  const headers = headerCount(rows);
  const unsafe = unsafeBoundaries(rows);
  const boundaries: TableBoundary[] = [];
  for (
    let rowIndex = Math.max(1, headers);
    rowIndex < rows.length;
    rowIndex += 1
  ) {
    const row = rows[rowIndex];
    if (row && !unsafe.has(rowIndex)) {
      boundaries.push({ pos: row.pos, offset: row.top });
    }
  }

  const repeatHeaderHeight = rows
    .slice(0, headers)
    .reduce((total, row) => total + row.height, 0);
  const firstBodyBoundary = boundaries.find(
    (boundary) => boundary.offset > repeatHeaderHeight + 0.5
  );
  const appliedHeight = appliedRows.reduce(
    (total, row) => total + row.getBoundingClientRect().height / scale,
    0
  );

  return {
    table: {
      boundaries,
      firstPageMinimum:
        firstBodyBoundary?.offset ??
        tableDom.getBoundingClientRect().height / scale - appliedHeight,
      repeatHeaderHeight,
      headerRows: rows.slice(0, headers).map((row) => row.pos),
      headerSignature: JSON.stringify(
        rows.slice(0, headers).map((row) => row.node.toJSON())
      ),
      columns: columnCount(tableNode),
    },
    appliedHeight,
  };
}
