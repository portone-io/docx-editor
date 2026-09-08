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
  type Plugin,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { editsShut } from "../../schema/protectionState";
import { panelPlugin } from "./panelState";

/** Where the menu should stand (viewport coordinates) */
export interface TableMenuAnchor {
  clientX: number;
  clientY: number;
}

/** The metadata is written from wherever the transaction was built, so a point is judged as one */
function isTableMenuAnchor(value: unknown): value is TableMenuAnchor {
  if (typeof value !== "object" || value === null) return false;
  const { clientX, clientY }: Partial<TableMenuAnchor> = value;
  return typeof clientX === "number" && typeof clientY === "number";
}

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
  const tr: Transaction = menu.opening(state.tr, anchor);
  if (!selectionCovers(state, cellPos)) {
    tr.setSelection(TextSelection.near(tr.doc.resolve(cellPos + 1)));
  }
  view.dispatch(tr);
}

const menu = panelPlugin<TableMenuAnchor>({
  name: "docxEditorTableMenu",
  isAnchor: isTableMenuAnchor,
  // Once the document changes, the cell the menu pointed at may no longer be there
  onDocChange: "close",
  props: {
    handleDOMEvents: {
      contextmenu(view, event) {
        // Where the body may not be edited there is nothing here to offer, so the click goes on
        // to the text menu or the browser
        if (editsShut(view.state)) return false;
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

/** The anchor of the currently open table menu. Null when it is closed */
export function tableMenuAnchor(state: EditorState): TableMenuAnchor | null {
  return menu.anchor(state);
}

/** Closes the menu. Called on running an item, Escape, clicking outside, and scrolling */
export const closeTableMenu: Command = menu.close;

export function tableContextMenu(): Plugin<TableMenuAnchor | null> {
  return menu.plugin;
}
