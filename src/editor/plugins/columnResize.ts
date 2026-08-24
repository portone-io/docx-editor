/**
 * Previews column-edge drags in the DOM and commits one dxa-based resize transaction on release.
 * Escape or an interrupted drag restores the rendered widths without editing the document.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";
import { spanCount } from "../../model/format";
import { editorClassNames } from "../../styles/classNames";
import { columnWidthPx } from "../../styles/inlineStyle";
import {
  buildResizeColumnTransaction,
  maxGridTotal,
  resizedGridCols,
} from "../../table/resize";
import {
  gridTotal,
  shiftedTableWidth,
  tableGridCols,
} from "../../table/widths";
import { documentGeometry } from "../documentStyles";
import { TableNodeView } from "../views/tableView";

/** Distance (px) that counts as being on an edge. Closer than this the cursor changes, and pressing starts a drag */
const EDGE_TOLERANCE_PX = 4;

/** Grid width per pixel used when the table could not be measured. 1pt = 20 dxa and 1pt on screen = 96/72px */
const FALLBACK_DXA_PER_PX = 15;

/** The grabbed edge. `col` is the grid column left of the edge, `side` is which side of the pressed cell that edge is */
export interface ColumnEdgeHit {
  col: number;
  side: "left" | "right";
}

/**
 * Which vertical edge of a cell this horizontal position points at.
 *
 * A merged cell only grabs the outer edges of the grid columns it covers.
 * The left end of the table is not grabbable because there is no column to move
 * (that edge means table indentation, which is a different thing).
 * We return the grabbed side as well so the guide line can stay glued to that side
 * during the drag.
 */
export function columnEdgeAt(
  clientX: number,
  cell: { readonly left: number; readonly right: number },
  covered: { readonly from: number; readonly to: number },
  tolerance: number = EDGE_TOLERANCE_PX
): ColumnEdgeHit | null {
  if (Math.abs(clientX - cell.right) <= tolerance) {
    return { col: covered.to - 1, side: "right" };
  }
  if (Math.abs(clientX - cell.left) <= tolerance && covered.from > 0) {
    return { col: covered.from - 1, side: "left" };
  }
  return null;
}

/** The range of grid columns covered by the cell the event happened in, plus its table */
interface CellSpot {
  tablePos: number;
  table: PMNode;
  from: number;
  to: number;
}

function cellSpotAt(view: EditorView, td: HTMLElement): CellSpot | null {
  const $pos = view.state.doc.resolve(view.posAtDOM(td, 0));
  // A cell's parent is a row and above that is the table, so the table is always two levels up
  for (let depth = $pos.depth; depth > 1; depth -= 1) {
    const cell = $pos.node(depth);
    if (cell.type.spec.tableRole !== "cell") continue;
    const table = $pos.node(depth - 2);
    if (table.type.spec.tableRole !== "table") return null;
    const tablePos = $pos.before(depth - 2);
    const from = TableMap.get(table).colCount(
      $pos.before(depth) - tablePos - 1
    );
    return { tablePos, table, from, to: from + spanCount(cell.attrs.colspan) };
  }
  return null;
}

/** One draggable edge. Holds both the document-side position and the screen-side element */
interface ColumnEdgeTarget extends CellSpot, ColumnEdgeHit {
  gridCols: number[];
  tableDom: HTMLTableElement;
  /** The cell whose edge was grabbed. This cell's side is exactly where the edge currently stands */
  cellDom: HTMLTableCellElement;
}

function columnElements(tableDom: HTMLTableElement): HTMLTableColElement[] {
  return Array.from(tableDom.querySelectorAll(":scope > colgroup > col"));
}

/** The column edge grabbable at this position. Null when not near an edge or when the table's grid is unknown */
function edgeUnder(
  view: EditorView,
  event: MouseEvent
): ColumnEdgeTarget | null {
  const node = event.target instanceof Node ? event.target : null;
  const element = node instanceof Element ? node : node?.parentElement;
  const td = element?.closest("td");
  const tableDom = td?.closest("table");
  if (!td || !tableDom || !view.dom.contains(td)) return null;

  const spot = cellSpotAt(view, td);
  if (!spot) return null;
  const gridCols = tableGridCols(spot.table);
  // If the grid and the table disagree on the column count, we cannot trust which column to move
  if (!gridCols || gridCols.length !== TableMap.get(spot.table).width) {
    return null;
  }

  const hit = columnEdgeAt(event.clientX, td.getBoundingClientRect(), spot);
  if (!hit || hit.col < 0 || hit.col >= gridCols.length) return null;
  return { ...spot, ...hit, gridCols, tableDom, cellDom: td };
}

/**
 * How far (in dxa) the grid should move for 1px of movement on screen.
 *
 * This is the ratio of grid width to rendered width. When the table is rendered exactly as
 * its grid says, it is always 15 (1pt = 20dxa = 96/72px); for tables whose width is a
 * percentage or whose grid is written only as proportions, it comes out to the value that
 * fits that table. Without this ratio, a tiny drag in a table with a small grid total would
 * make the width jump straight to its limit.
 */
function dxaPerPx(gridCols: number[], tableDom: HTMLElement): number {
  const drawn = tableDom.getBoundingClientRect().width;
  return drawn > 0 ? gridTotal(gridCols) / drawn : FALLBACK_DXA_PER_PX;
}

/** What we hold on to during a drag. The preview only touches the DOM recorded here */
interface Drag {
  readonly view: EditorView;
  readonly target: ColumnEdgeTarget;
  readonly maxTotal: number;
  readonly startX: number;
  /** Grid width per pixel, measured when the drag started. It does not change during the drag */
  readonly dxaPerPx: number;
  readonly cols: readonly HTMLTableColElement[];
  /** The original inline widths, used to revert the preview */
  readonly colWidths: readonly string[];
  readonly tableWidth: string;
  /** The vertical line showing where the edge will land */
  readonly guide: HTMLElement;
  /** The width (dxa) that the current drag position asks for on the column left of the edge */
  width: number;
}

/**
 * The vertical line showing where the edge will land.
 *
 * It is positioned fixed to the viewport so the page cannot clip it.
 * Since it is not editable content it lives outside the document DOM (on body), and it is
 * removed once the drag ends.
 */
function createGuide(): HTMLElement {
  const guide = document.createElement("div");
  guide.className = editorClassNames.columnResizeGuide;
  document.body.appendChild(guide);
  return guide;
}

/**
 * Moves the guide line to where the edge currently stands.
 *
 * The position is read straight off the side of the grabbed cell. That way, even at the
 * moment the width stops because it hit a limit, the line points at the real edge rather
 * than at the mouse.
 */
function moveGuide(drag: Drag): void {
  const { cellDom, side, tableDom } = drag.target;
  const table = tableDom.getBoundingClientRect();
  drag.guide.style.left = `${cellDom.getBoundingClientRect()[side]}px`;
  drag.guide.style.top = `${table.top}px`;
  drag.guide.style.height = `${table.height}px`;
}

function showPreview(drag: Drag, gridCols: number[]): void {
  gridCols.forEach((dxa, at) => {
    const col = drag.cols[at];
    if (col) col.style.width = `${columnWidthPx(dxa)}px`;
  });
  // When the table width is pinned, the fixed layout redistributes the columns within that width.
  // For the table to actually grow while dragging the last edge, the table width has to move too.
  const width = shiftedTableWidth(drag.target.table, gridCols);
  if (width?.type === "dxa") {
    drag.target.tableDom.style.width = `${columnWidthPx(width.twips)}px`;
  }
  moveGuide(drag);
}

function hidePreview(drag: Drag): void {
  drag.cols.forEach((col, at) => {
    col.style.width = drag.colWidths[at] ?? "";
  });
  drag.target.tableDom.style.width = drag.tableWidth;
  drag.guide.remove();
}

export function columnResize(): Plugin {
  let drag: Drag | null = null;

  function stop(): void {
    if (!drag) return;
    window.removeEventListener("mousemove", onMove, true);
    window.removeEventListener("mouseup", onUp, true);
    window.removeEventListener("keydown", onKey, true);
    hidePreview(drag);
    drag.view.dom.classList.remove(editorClassNames.columnResize);
    drag = null;
  }

  function onMove(event: MouseEvent): void {
    if (!drag) return;
    event.preventDefault();
    const { gridCols, col } = drag.target;
    const moved = (event.clientX - drag.startX) * drag.dxaPerPx;
    drag.width = gridCols[col] + Math.round(moved);
    const preview = resizedGridCols(gridCols, col, drag.width, drag.maxTotal);
    showPreview(drag, preview ?? gridCols);
  }

  function onUp(): void {
    const current = drag;
    if (!current) return;
    // The commit transaction re-renders the table, so the preview is torn down first
    stop();
    const tr = buildResizeColumnTransaction(
      current.view.state,
      {
        tablePos: current.target.tablePos,
        col: current.target.col,
        width: current.width,
      },
      documentGeometry(current.view.state)
    );
    if (tr) current.view.dispatch(tr);
  }

  function onKey(event: KeyboardEvent): void {
    if (event.key !== "Escape") return;
    event.preventDefault();
    stop();
  }

  function start(view: EditorView, event: MouseEvent): boolean {
    const target = edgeUnder(view, event);
    if (!target) return false;
    const cols = columnElements(target.tableDom);
    if (cols.length !== target.gridCols.length) return false;

    drag = {
      view,
      target,
      maxTotal: maxGridTotal(target.table, documentGeometry(view.state)),
      startX: event.clientX,
      dxaPerPx: dxaPerPx(target.gridCols, target.tableDom),
      cols,
      colWidths: cols.map((col) => col.style.width),
      tableWidth: target.tableDom.style.width,
      guide: createGuide(),
      width: target.gridCols[target.col],
    };
    moveGuide(drag);
    view.dom.classList.add(editorClassNames.columnResize);
    window.addEventListener("mousemove", onMove, true);
    window.addEventListener("mouseup", onUp, true);
    window.addEventListener("keydown", onKey, true);
    return true;
  }

  return new Plugin({
    view: () => ({ destroy: stop }),
    props: {
      // Keeping ProseMirror from reverting the colgroup the preview edited is all this view does
      nodeViews: { table: (node) => new TableNodeView(node) },
      handleDOMEvents: {
        mousemove(view, event) {
          if (drag) return false;
          const over = view.editable && edgeUnder(view, event) !== null;
          view.dom.classList.toggle(editorClassNames.columnResize, over);
          return false;
        },
        mouseleave(view) {
          if (!drag) view.dom.classList.remove(editorClassNames.columnResize);
          return false;
        },
        mousedown(view, event) {
          if (!view.editable || drag || event.button !== 0) return false;
          if (!start(view, event)) return false;
          // A press that grabs an edge starts neither a text selection nor a cell selection
          event.preventDefault();
          return true;
        },
      },
    },
  });
}
