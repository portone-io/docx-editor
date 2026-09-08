/**
 * The elements a reading takes as blocks of their own.
 *
 * Everything else in a piece of clipboard HTML is read by the shape of the markup alone: an
 * element either names a block or it is text in the paragraph around it. A block reader is how an
 * element that means more than its shape says - a table, a Word paragraph that is really a list
 * item - is read as what it means, without the reading itself learning about the applications
 * one by one.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { GridRect } from "../../docx/tableFormatting";
import {
  createTableNodeFrom,
  isTableSide,
  MAX_TABLE_SIDE,
  type TableCellPlan,
} from "../../docx/tableTemplate";
import { type ListKind, MAX_ILVL } from "../../numbering/listTemplate";
import { contextFor, type InlineContext } from "./inlineFormatting";
import type { HtmlReadContext } from "./readContext";

/** Where one paragraph being read stands in a list */
export interface ListPlacement {
  kind: ListKind;
  /** The number the list wears here, or null where this document may start no list */
  numId: number | null;
  /** How deep the item sits, counted from zero */
  level: number;
}

/** What a block reader may ask of the reading it is part of */
export interface HtmlBlockHost {
  /** What the markup is being read against, including the application that wrote it */
  readonly context: HtmlReadContext;
  /** Whether the reading already stands inside a cell of a table it is reading */
  readonly inTable: boolean;
  /** The blocks the content of one cell reads as. None where the cell holds nothing */
  readCell(cell: ParentNode, inline: InlineContext): readonly PMNode[];
  /** The inline nodes these nodes read as */
  readInline(nodes: Iterable<Node>, inline: InlineContext): readonly PMNode[];
  /** One paragraph of the reading, as an item of the list it belongs to when it is one */
  paragraph(content: readonly PMNode[], list?: ListPlacement | null): PMNode;
  /**
   * The number the pasted list this key names takes here. The same key answers the same number,
   * so the paragraphs of one list join it rather than each starting a list of their own.
   */
  listNumber(key: string, kind: ListKind): number | null;
}

/**
 * One kind of element read as blocks of its own rather than as text where it stands.
 *
 * A reader answers null for an element it has nothing to say about, so that deciding whether an
 * element is its own and reading it are one pass rather than two.
 */
export interface HtmlBlockReader {
  read(
    element: HTMLElement,
    inline: InlineContext,
    host: HtmlBlockHost
  ): readonly PMNode[] | null;
}

const CELL_TAGS = new Set(["TD", "TH"]);
const ROW_GROUP_TAGS = new Set(["THEAD", "TBODY", "TFOOT"]);

/** One cell of an HTML table and the block of the grid it covers */
interface PlacedCell {
  element: HTMLElement;
  rect: GridRect;
}

/**
 * The rows of this table alone. A browser puts the rows a writer left bare into a `<tbody>` of its
 * own, and a table standing inside a cell belongs to that cell rather than to this grid.
 */
function rowsOf(table: HTMLElement): HTMLElement[] {
  return [...table.children].flatMap((child) => {
    if (!(child instanceof HTMLElement)) return [];
    if (child.tagName === "TR") return [child];
    if (!ROW_GROUP_TAGS.has(child.tagName)) return [];
    return [...child.children].flatMap((row) =>
      row instanceof HTMLElement && row.tagName === "TR" ? [row] : []
    );
  });
}

function cellsOf(row: HTMLElement): HTMLElement[] {
  return [...row.children].flatMap((cell) =>
    cell instanceof HTMLElement && CELL_TAGS.has(cell.tagName) ? [cell] : []
  );
}

/** How far a cell reaches. A span that is missing, unreadable or absurd reaches one cell */
function spanOf(
  element: HTMLElement,
  name: "colspan" | "rowspan",
  limit: number
): number {
  const value = Number.parseInt(element.getAttribute(name) ?? "", 10);
  return Number.isFinite(value) && value > 1 ? Math.min(value, limit) : 1;
}

/**
 * Where every cell of the table sits.
 *
 * A cell reaching down rows leaves the columns it covers taken on the rows below, which is what
 * decides where the cells after it begin, and the widest row is what the grid is as wide as. A row
 * holding no cell of its own is left out rather than written as an empty row, since the model has
 * a merge standing only where it starts. Null where the table is too large to model.
 */
function gridOf(table: HTMLElement): readonly (readonly PlacedCell[])[] | null {
  const rows = rowsOf(table).filter((row) => cellsOf(row).length > 0);
  if (!isTableSide(rows.length)) return null;
  const covered = rows.map(() => new Set<number>());
  const cells: PlacedCell[][] = rows.map(() => []);
  let cols = 0;
  rows.forEach((row, top) => {
    let left = 0;
    for (const element of cellsOf(row)) {
      while (covered[top]?.has(left)) left += 1;
      const colspan = spanOf(element, "colspan", MAX_TABLE_SIDE);
      const rowspan = Math.min(
        spanOf(element, "rowspan", MAX_TABLE_SIDE),
        rows.length - top
      );
      for (let covering = top; covering < top + rowspan; covering += 1) {
        for (let column = left; column < left + colspan; column += 1) {
          covered[covering]?.add(column);
        }
      }
      cells[top]?.push({
        element,
        rect: { top, bottom: top + rowspan, left, right: left + colspan },
      });
      left += colspan;
      cols = Math.max(cols, left);
    }
  });
  return isTableSide(cols) ? cells : null;
}

/**
 * An HTML table, read as a table of this document.
 *
 * The cells hold what their own markup reads as, one paragraph per block they hold, and the table
 * is as wide as a table inserted here would be: a pasted table is drawn and written out exactly
 * like one the toolbar put there, since nothing of the width a foreign document measured its own
 * paper against would mean the same here.
 *
 * A table inside a cell is not read as a table: the model only makes the outer one editable, so
 * the inner one is left to the reading around it, which takes its text into the cell.
 */
export const tableBlockReader: HtmlBlockReader = {
  read: (element, inline, host) => {
    if (element.tagName !== "TABLE" || host.inTable) return null;
    const grid = gridOf(element);
    if (grid === null) return null;
    const inside = contextFor(inline, element);
    const cells: readonly (readonly TableCellPlan[])[] = grid.map((row) =>
      row.map(({ element: cell, rect }) => ({
        rect,
        content: host.readCell(cell, contextFor(inside, cell)),
      }))
    );
    return [createTableNodeFrom(cells, host.context.geometry)];
  },
};

/**
 * How Word says a paragraph is an item of a list: which list definition it belongs to, how deep it
 * sits, and which list of the document it is one of (`mso-list:l0 level1 lfo1`).
 */
const WORD_LIST =
  /mso-list:\s*l(?<list>\d+)\s+level(?<level>\d+)\s+lfo(?<applied>\d+)/i;

/** The marker Word draws for a reader that numbers no lists itself */
const IGNORED_MARKER = /mso-list:\s*Ignore/i;
const MARKER_START = /^\[if\s+!supportLists\]/i;
const MARKER_END = /^\[endif\]/i;

/** A marker counting off items rather than standing in front of each of them */
const COUNTED_MARKER = /^\s*(?:\d+|[A-Za-z]+)\s*[.)]/;

/** Where a Word paragraph says it stands in a list */
interface WordListItem {
  /** The list definition and the list of the document, which together name one list */
  key: string;
  level: number;
}

function wordListItemOf(element: HTMLElement): WordListItem | null {
  const found = WORD_LIST.exec(element.getAttribute("style") ?? "")?.groups;
  if (found === undefined) return null;
  const level = Number.parseInt(found.level ?? "", 10);
  return {
    key: `${found.list}/${found.applied}`,
    // Word counts its levels from one, and no document has more levels than the model holds
    level: Number.isFinite(level)
      ? Math.min(Math.max(level, 1), MAX_ILVL + 1) - 1
      : 0,
  };
}

/**
 * The content of a Word list paragraph with the marker Word drew left out, and that marker.
 *
 * Word writes the bullet or the number into the paragraph itself, between two conditional comments
 * or in a span it marks to be ignored, so that an application numbering no lists still shows one.
 * Here the list carries its own numbering, so the drawn marker would be a second one.
 */
function withoutMarker(element: HTMLElement): {
  marker: string;
  content: readonly Node[];
} {
  const content: Node[] = [];
  let marker = "";
  let inMarker = false;
  for (const child of element.childNodes) {
    if (child.nodeType === child.COMMENT_NODE) {
      const data = child.nodeValue ?? "";
      if (MARKER_START.test(data)) inMarker = true;
      else if (MARKER_END.test(data)) inMarker = false;
      continue;
    }
    const ignored =
      child instanceof HTMLElement &&
      IGNORED_MARKER.test(child.getAttribute("style") ?? "");
    if (inMarker || ignored) marker += child.textContent ?? "";
    else content.push(child);
  }
  return { marker, content };
}

/**
 * A Word paragraph that is an item of a list, read as an item of a list here.
 *
 * Word puts no `<ul>` or `<ol>` on the clipboard: every item is a paragraph saying which list it
 * belongs to, so the paragraphs of one list are joined by that name rather than by standing inside
 * one element. Whether the list counts its items is only visible in the marker Word drew.
 */
export const wordListReader: HtmlBlockReader = {
  read: (element, inline, host) => {
    if (host.context.source !== "word") return null;
    const item = wordListItemOf(element);
    if (item === null) return null;
    const { marker, content } = withoutMarker(element);
    const kind: ListKind = COUNTED_MARKER.test(marker) ? "numbered" : "bullet";
    return [
      host.paragraph(host.readInline(content, contextFor(inline, element)), {
        kind,
        numId: host.listNumber(item.key, kind),
        level: item.level,
      }),
    ];
  },
};

/** The readers every reading consults, in the order they are tried */
export const DEFAULT_BLOCK_READERS: readonly HtmlBlockReader[] = [
  tableBlockReader,
  wordListReader,
];
