/**
 * The kind of a table: it is parted between rows, and a continued page opens a spacer row and
 * repeats the headers the table opens with.
 *
 * Both sides read the table as it stands on screen without treating the rows this very engine
 * drew as content: the measurement takes their heights back off, and the projection is derived
 * again from the document, so an edit to a header row redraws it with nothing measured anew.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Decoration, type EditorView } from "prosemirror-view";
import { toRowFormat } from "../../model/format";
import { editorAttributes } from "../../styles/classNames";
import type { BlockKind, BreakCandidate } from "../blockKinds";

const PAGE_BREAK_BR = `br[${editorAttributes.breakType}="page"]`;

interface RowEntry {
  node: PMNode;
  pos: number;
  top: number;
  height: number;
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
 * with. Read off the table itself so the projection follows an edit to a header row.
 */
function headerRowsOf(tableNode: PMNode, tablePos: number): readonly number[] {
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

function tableSpace(height: number, columns: number): HTMLElement {
  const row = document.createElement("tr");
  row.setAttribute(editorAttributes.tablePageSpace, `${height}`);
  row.setAttribute("aria-hidden", "true");
  row.setAttribute("contenteditable", "false");
  row.style.height = `${height}px`;

  const cell = document.createElement("td");
  cell.colSpan = columns;
  cell.style.height = `${height}px`;
  row.append(cell);
  return row;
}

function repeatedHeader(view: EditorView, rowPos: number): HTMLElement {
  const source = view.nodeDOM(rowPos);
  const row =
    source instanceof HTMLElement && source.tagName === "TR"
      ? (source.cloneNode(true) as HTMLElement)
      : document.createElement("tr");
  row.setAttribute(editorAttributes.tableRepeatedHeader, "");
  row.setAttribute("aria-hidden", "true");
  row.setAttribute("contenteditable", "false");
  row.removeAttribute("id");
  row.querySelectorAll("[id]").forEach((element) => {
    element.removeAttribute("id");
  });
  return row;
}

export const tableKind: BlockKind = {
  name: "table",

  matches: (node) => node.type.spec.tableRole === "table",

  measure({ view, node, pos, dom, scale }) {
    const appliedRows = paginationRows(dom);
    const rows = rowEntries(view, node, pos, dom, appliedRows, scale);
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
      // The smallest useful first piece: the headers followed by one body row group
      minFirstPiece:
        firstBodyBoundary?.offset ??
        dom.getBoundingClientRect().height / scale - appliedHeight,
      // A space cannot be opened inside a cell, so a break standing in one is answered after the
      // whole table
      breakAfter: dom.querySelectorAll(PAGE_BREAK_BR).length > 0,
      appliedHeight,
    };
  },

  holdsCut: (doc, at) => doc.nodeAt(at)?.type.spec.tableRole === "row",

  decorate(pos, node, cuts, into) {
    if (cuts.length === 0) return;
    const headerRows = headerRowsOf(node, pos);
    const columns = columnCount(node);
    for (const cut of cuts) {
      into.push(
        Decoration.widget(cut.at, () => tableSpace(cut.height, columns), {
          key: `table-page-space-${cut.at}-${cut.height}-${columns}`,
          side: -100,
        })
      );
      headerRows.forEach((rowPos, headerIndex) => {
        into.push(
          // No `key`: a keyed widget is held to be the same one and left alone, and this one is a
          // copy of a row that the very edit rebuilding these decorations may have just changed
          Decoration.widget(cut.at, (view) => repeatedHeader(view, rowPos), {
            side: -90 + headerIndex,
          })
        );
      });
    }
  },
};
