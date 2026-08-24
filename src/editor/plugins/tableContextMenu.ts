/**
 * Right-clicking in a table cell brings up our own menu.
 *
 * This file does two things: block the browser's default menu, and report where the menu should
 * stand. Drawing the menu is the React side's job.
 *
 * Only the click that asks for the row and column actions reaches this plugin: a click on
 * selected text is taken first by the text menu (`editor/textContextMenu`), which stands ahead
 * of this one and hands back a click with nothing selected inside a cell.
 */

import type { ResolvedPos } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  Plugin,
  PluginKey,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

/** Where the menu should stand (viewport coordinates) */
export interface TableMenuAnchor {
  clientX: number;
  clientY: number;
}

/** The signal to close the menu. A single fixed value, so it cannot be confused with anchor data */
const CLOSE = "close";

const menuKey = new PluginKey<TableMenuAnchor | null>("docxEditorTableMenu");

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function toAnchor(value: unknown): TableMenuAnchor | null {
  if (!isRecord(value)) return null;
  const { clientX, clientY } = value;
  if (typeof clientX !== "number" || typeof clientY !== "number") return null;
  return { clientX, clientY };
}

/** The anchor of the currently open table menu. Null when it is closed */
export function tableMenuAnchor(state: EditorState): TableMenuAnchor | null {
  return menuKey.getState(state) ?? null;
}

/** Closes the menu. Called on running an item, Escape, clicking outside, and scrolling */
export const closeTableMenu: Command = (state, dispatch) => {
  if (tableMenuAnchor(state) === null) return false;
  if (dispatch) dispatch(state.tr.setMeta(menuKey, CLOSE));
  return true;
};

/** The start position of the table cell containing this position. Null when outside a cell */
function cellStart($pos: ResolvedPos): number | null {
  for (let depth = $pos.depth; depth > 0; depth -= 1) {
    if ($pos.node(depth).type.spec.tableRole === "cell") {
      return $pos.before(depth);
    }
  }
  return null;
}

/** The start position of the table cell the event happened in. Null when outside a cell */
function cellAtEvent(
  view: EditorView,
  target: EventTarget | null
): number | null {
  const node = target instanceof Node ? target : null;
  const element = node instanceof Element ? node : node?.parentElement;
  const cell = element?.closest("td");
  if (!cell || !view.dom.contains(cell)) return null;
  return cellStart(view.state.doc.resolve(view.posAtDOM(cell, 0)));
}

/** Whether this cell is inside the current selection. A selection covering several cells is preserved as is */
function selectionCovers(state: EditorState, cellPos: number): boolean {
  return state.selection.ranges.some(
    (range) => cellStart(range.$from) === cellPos
  );
}

/**
 * Opens the menu, moving the selection to the right-clicked cell if needed.
 * What the user expects to act on is always the cell they just clicked.
 */
function openMenu(
  view: EditorView,
  cellPos: number,
  anchor: TableMenuAnchor
): void {
  const state = view.state;
  const tr: Transaction = state.tr.setMeta(menuKey, anchor);
  if (!selectionCovers(state, cellPos)) {
    tr.setSelection(TextSelection.near(tr.doc.resolve(cellPos + 1)));
  }
  view.dispatch(tr);
}

export function tableContextMenu(): Plugin<TableMenuAnchor | null> {
  return new Plugin<TableMenuAnchor | null>({
    key: menuKey,
    state: {
      init: () => null,
      apply: (tr, current) => {
        const meta: unknown = tr.getMeta(menuKey);
        if (meta === CLOSE) return null;
        const anchor = toAnchor(meta);
        if (anchor) return anchor;
        // Once the document changes, the cell the menu pointed at may no longer be there
        return tr.docChanged ? null : current;
      },
    },
    props: {
      handleDOMEvents: {
        contextmenu(view, event) {
          // In read-only mode there is nothing to edit, so the browser menu is the better choice
          if (!view.editable) return false;
          const cellPos = cellAtEvent(view, event.target);
          if (cellPos === null) return false;
          event.preventDefault();
          openMenu(view, cellPos, {
            clientX: event.clientX,
            clientY: event.clientY,
          });
          return true;
        },
      },
    },
  });
}
