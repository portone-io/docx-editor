import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { isLockedCell } from "../../schema/locks";
import { editorClassNames } from "../../styles/classNames";
import {
  buildResizeRowTransaction,
  resizedRowHeight,
  rowPositionAt,
} from "../../table/rowResize";

const EDGE_TOLERANCE_PX = 4;
const rowResizeKey = new PluginKey<DecorationSet>("docxRowResize");

export interface RowEdgeHit {
  row: number;
  side: "top" | "bottom";
}

/** The row boundary under a pointer, including the end of a vertically merged cell. */
export function rowEdgeAt(
  clientY: number,
  cell: { readonly top: number; readonly bottom: number },
  covered: { readonly from: number; readonly to: number },
  tolerance: number = EDGE_TOLERANCE_PX
): RowEdgeHit | null {
  if (Math.abs(clientY - cell.bottom) <= tolerance) {
    return { row: covered.to - 1, side: "bottom" };
  }
  if (Math.abs(clientY - cell.top) <= tolerance && covered.from > 0) {
    return { row: covered.from - 1, side: "top" };
  }
  return null;
}

interface CellSpot {
  tablePos: number;
  table: PMNode;
  from: number;
  to: number;
}

function cellSpotAt(view: EditorView, td: HTMLElement): CellSpot | null {
  const $pos = view.state.doc.resolve(view.posAtDOM(td, 0));
  for (let depth = $pos.depth; depth > 1; depth -= 1) {
    const cell = $pos.node(depth);
    if (cell.type.spec.tableRole !== "cell") continue;
    const table = $pos.node(depth - 2);
    if (table.type.spec.tableRole !== "table") return null;
    const tablePos = $pos.before(depth - 2);
    const rect = TableMap.get(table).findCell(
      $pos.before(depth) - tablePos - 1
    );
    return { tablePos, table, from: rect.top, to: rect.bottom };
  }
  return null;
}

interface RowEdgeTarget extends CellSpot, RowEdgeHit {
  tableDom: HTMLTableElement;
  rowDom: HTMLTableRowElement;
  rowPos: number;
}

function edgeUnder(view: EditorView, event: MouseEvent): RowEdgeTarget | null {
  const node = event.target instanceof Node ? event.target : null;
  const element = node instanceof Element ? node : node?.parentElement;
  const td = element?.closest("td");
  const tableDom = td?.closest("table");
  if (!td || !tableDom || !view.dom.contains(td)) return null;
  const spot = cellSpotAt(view, td);
  if (!spot) return null;
  const hit = rowEdgeAt(event.clientY, td.getBoundingClientRect(), spot);
  if (!hit || hit.row < 0 || hit.row >= spot.table.childCount) return null;
  const row = spot.table.child(hit.row);
  if (row.children.some(isLockedCell)) return null;
  const rowPos = rowPositionAt(spot.tablePos, spot.table, hit.row);
  const rowDom = view.nodeDOM(rowPos);
  if (!(rowDom instanceof HTMLTableRowElement)) return null;
  return { ...spot, ...hit, tableDom, rowDom, rowPos };
}

function visualScale(view: EditorView): number {
  const layer = view.dom.closest(`.${editorClassNames.pageLayer}`);
  const raw = layer instanceof HTMLElement ? getComputedStyle(layer).zoom : "";
  const parsed = Number.parseFloat(raw);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 1;
}

interface Drag {
  readonly view: EditorView;
  readonly target: RowEdgeTarget;
  readonly startY: number;
  readonly startHeightPt: number;
  readonly scale: number;
  readonly guide: HTMLElement;
  heightPt: number;
  moved: boolean;
}

interface RowPreview {
  rowPos: number;
  heightPt: number;
}

function previewDecorations(doc: PMNode, preview: RowPreview): DecorationSet {
  const row = doc.nodeAt(preview.rowPos);
  if (row?.type.spec.tableRole !== "row") return DecorationSet.empty;
  return DecorationSet.create(doc, [
    Decoration.node(preview.rowPos, preview.rowPos + row.nodeSize, {
      style: `height:${preview.heightPt}pt !important`,
    }),
  ]);
}

function createGuide(): HTMLElement {
  const guide = document.createElement("div");
  guide.className = editorClassNames.rowResizeGuide;
  document.body.appendChild(guide);
  return guide;
}

function moveGuide(drag: Drag): void {
  const table = drag.target.tableDom.getBoundingClientRect();
  drag.guide.style.left = `${table.left}px`;
  drag.guide.style.top = `${drag.target.rowDom.getBoundingClientRect().bottom}px`;
  drag.guide.style.width = `${table.width}px`;
}

function showPreview(drag: Drag): void {
  drag.view.dispatch(
    drag.view.state.tr
      .setMeta(rowResizeKey, {
        rowPos: drag.target.rowPos,
        heightPt: drag.heightPt,
      } satisfies RowPreview)
      .setMeta("addToHistory", false)
  );
  moveGuide(drag);
}

function hidePreview(drag: Drag): void {
  drag.view.dispatch(
    drag.view.state.tr
      .setMeta(rowResizeKey, null)
      .setMeta("addToHistory", false)
  );
  drag.guide.remove();
}

export function rowResize(): Plugin<DecorationSet> {
  let drag: Drag | null = null;

  function stop(clearPreview = true): void {
    if (!drag) return;
    window.removeEventListener("mousemove", onMove, true);
    window.removeEventListener("mouseup", onUp, true);
    window.removeEventListener("keydown", onKey, true);
    if (clearPreview) hidePreview(drag);
    else drag.guide.remove();
    drag.view.dom.classList.remove(editorClassNames.rowResize);
    drag = null;
  }

  function onMove(event: MouseEvent): void {
    if (!drag) return;
    event.preventDefault();
    drag.moved ||= event.clientY !== drag.startY;
    drag.heightPt = resizedRowHeight(
      drag.startHeightPt,
      event.clientY - drag.startY,
      drag.scale
    );
    showPreview(drag);
  }

  function onUp(): void {
    const current = drag;
    if (!current) return;
    stop();
    if (!current.moved) return;
    const tr = buildResizeRowTransaction(current.view.state, {
      tablePos: current.target.tablePos,
      row: current.target.row,
      heightPt: current.heightPt,
    });
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
    const scale = visualScale(view);
    const startHeightPt =
      (target.rowDom.getBoundingClientRect().height / scale) * (72 / 96);
    drag = {
      view,
      target,
      startY: event.clientY,
      startHeightPt,
      scale,
      guide: createGuide(),
      heightPt: startHeightPt,
      moved: false,
    };
    moveGuide(drag);
    view.dom.classList.add(editorClassNames.rowResize);
    window.addEventListener("mousemove", onMove, true);
    window.addEventListener("mouseup", onUp, true);
    window.addEventListener("keydown", onKey, true);
    return true;
  }

  return new Plugin({
    key: rowResizeKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, decorations) {
        const preview = tr.getMeta(rowResizeKey) as
          | RowPreview
          | null
          | undefined;
        if (preview === undefined) return decorations.map(tr.mapping, tr.doc);
        return preview === null
          ? DecorationSet.empty
          : previewDecorations(tr.doc, preview);
      },
    },
    view: () => ({ destroy: () => stop(false) }),
    props: {
      decorations: (state) => rowResizeKey.getState(state),
      handleDOMEvents: {
        mousemove(view, event) {
          if (drag) return false;
          const over = view.editable && edgeUnder(view, event) !== null;
          view.dom.classList.toggle(editorClassNames.rowResize, over);
          return false;
        },
        mouseleave(view) {
          if (!drag) view.dom.classList.remove(editorClassNames.rowResize);
          return false;
        },
        mousedown(view, event) {
          if (!view.editable || drag || event.button !== 0) return false;
          if (!start(view, event)) return false;
          event.preventDefault();
          return true;
        },
      },
    },
  });
}
