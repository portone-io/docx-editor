/** Measures safe row boundaries without treating pagination decorations as document content. */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { toRowFormat } from "../model/format";
import { editorAttributes } from "../styles/classNames";
import type { BreakCandidate } from "./blockKinds";

interface RowEntry {
  node: PMNode;
  pos: number;
  top: number;
  height: number;
}

export interface TableMeasure {
  /** Every row that may start the continued part of the table */
  candidates: readonly BreakCandidate[];
  /** The smallest useful first piece: headers followed by one body row group */
  minFirstPiece: number;
  /** Height already added by the previous pagination pass */
  appliedHeight: number;
  /** Document positions of the contiguous header rows at the start of the table */
  headerRows: readonly number[];
  /** Changes when any projected header content or formatting changes */
  headerSignature: string;
  columns: number;
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

function repeatsHeader(row: PMNode | undefined): boolean {
  return toRowFormat(row?.attrs.format)?.repeatHeader === true;
}

function headerCount(rows: readonly RowEntry[]): number {
  let count = 0;
  while (count < rows.length && repeatsHeader(rows[count]?.node)) {
    count += 1;
  }
  return count;
}

/**
 * Where the rows a continued page repeats stand: the run of `repeatHeader` rows the table opens
 * with. Read off the table itself so the projection follows an edit to a header row
 * (`page/pageDecorations`).
 */
export function headerRowsOf(
  tableNode: PMNode,
  tablePos: number
): readonly number[] {
  const found: number[] = [];
  let offset = 0;
  for (let index = 0; index < tableNode.childCount; index += 1) {
    const row = tableNode.child(index);
    if (!repeatsHeader(row)) break;
    found.push(tablePos + 1 + offset);
    offset += row.nodeSize;
  }
  return found;
}

/** How many grid columns a row of this table spans, so a spacer row can cover it */
export function columnCount(tableNode: PMNode): number {
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
  const headerRows = rows.slice(0, headers);
  const repeatHeaderHeight = headerRows.reduce(
    (total, row) => total + row.height,
    0
  );

  // A row a page may start on carries the repeated headers onto that page ahead of itself, and
  // is never a cut of its own: it is taken only where the rows after it would run off the page
  const candidates: BreakCandidate[] = [];
  for (
    let rowIndex = Math.max(1, headers);
    rowIndex < rows.length;
    rowIndex += 1
  ) {
    const row = rows[rowIndex];
    if (row && !unsafe.has(rowIndex)) {
      candidates.push({
        at: row.pos,
        offset: row.top,
        forced: false,
        repeatHeight: repeatHeaderHeight,
      });
    }
  }

  const firstBodyBoundary = candidates.find(
    (candidate) => candidate.offset > repeatHeaderHeight + 0.5
  );
  const appliedHeight = appliedRows.reduce(
    (total, row) => total + row.getBoundingClientRect().height / scale,
    0
  );

  return {
    candidates,
    minFirstPiece:
      firstBodyBoundary?.offset ??
      tableDom.getBoundingClientRect().height / scale - appliedHeight,
    appliedHeight,
    headerRows: headerRows.map((row) => row.pos),
    headerSignature: JSON.stringify(headerRows.map((row) => row.node.toJSON())),
    columns: columnCount(tableNode),
  };
}
