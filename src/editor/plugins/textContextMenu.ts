/**
 * Replaces the native text context menu and moves the selection to the clicked position. Unselected
 * table cells are left to the table context-menu plugin.
 */

import {
  type Command,
  type EditorState,
  Plugin,
  PluginKey,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";

/** Where the menu should stand (viewport coordinates) */
export interface TextMenuAnchor {
  clientX: number;
  clientY: number;
}

/** The signal to close the menu. A single fixed value, so it cannot be confused with anchor data */
const CLOSE = "close";

const menuKey = new PluginKey<TextMenuAnchor | null>("docxEditorTextMenu");

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function toAnchor(value: unknown): TextMenuAnchor | null {
  if (!isRecord(value)) return null;
  const { clientX, clientY } = value;
  if (typeof clientX !== "number" || typeof clientY !== "number") return null;
  return { clientX, clientY };
}

/** The anchor of the currently open text menu. Null when it is closed */
export function textMenuAnchor(state: EditorState): TextMenuAnchor | null {
  return menuKey.getState(state) ?? null;
}

/** Closes the menu. Called on running an item, Escape, clicking outside, and scrolling */
export const closeTextMenu: Command = (state, dispatch) => {
  if (textMenuAnchor(state) === null) return false;
  if (dispatch) dispatch(state.tr.setMeta(menuKey, CLOSE));
  return true;
};

function isInEditor(view: EditorView, target: EventTarget | null): boolean {
  return target instanceof Node && view.dom.contains(target);
}

/** Whether the click landed in a table cell of this editor */
function isInTableCell(view: EditorView, target: EventTarget | null): boolean {
  const node = target instanceof Node ? target : null;
  const element = node instanceof Element ? node : node?.parentElement;
  const cell = element?.closest("td");
  return !!cell && view.dom.contains(cell);
}

/**
 * The spot in the document the click landed on.
 * Where the browser cannot say, the caret the state holds stands in, so the click counts as one
 * that moved nothing.
 */
function clickedSpot(view: EditorView, event: MouseEvent): number {
  const found = view.posAtCoords({ left: event.clientX, top: event.clientY });
  return found?.pos ?? view.state.selection.head;
}

/** Whether this spot lies within the selected text. An empty selection covers no spot at all */
function isInSelection(state: EditorState, spot: number): boolean {
  const selection = state.selection;
  return (
    !selection.empty &&
    selection.ranges.some(
      (range) => range.$from.pos <= spot && spot <= range.$to.pos
    )
  );
}

/**
 * Whether this click is the one the table menu is for.
 * Either it landed outside the selected text, so there is no stretch of text it is about, or
 * whole cells are selected: both are about the table, and a block of selected cells is how
 * merging is reached.
 */
function forTableMenu(
  view: EditorView,
  target: EventTarget | null,
  spot: number
): boolean {
  const aboutCells =
    view.state.selection instanceof CellSelection ||
    !isInSelection(view.state, spot);
  return aboutCells && isInTableCell(view, target);
}

/** Opens the menu, taking the caret along to the clicked spot when that spot is not selected */
function openMenu(
  state: EditorState,
  event: MouseEvent,
  spot: number
): Transaction {
  const tr = state.tr.setMeta(menuKey, {
    clientX: event.clientX,
    clientY: event.clientY,
  });
  if (isInSelection(state, spot)) return tr;
  return tr.setSelection(TextSelection.near(tr.doc.resolve(spot)));
}

export function textContextMenu(): Plugin<TextMenuAnchor | null> {
  return new Plugin<TextMenuAnchor | null>({
    key: menuKey,
    state: {
      init: () => null,
      apply: (tr, current) => {
        const meta: unknown = tr.getMeta(menuKey);
        if (meta === CLOSE) return null;
        const anchor = toAnchor(meta);
        if (anchor) return anchor;
        // Once the document changes, what the menu was aimed at may no longer be there
        return tr.docChanged ? null : current;
      },
    },
    props: {
      handleDOMEvents: {
        contextmenu(view, event) {
          // In read-only mode there is nothing to edit, so the browser menu is the better choice
          if (!view.editable) return false;
          if (!isInEditor(view, event.target)) return false;
          const spot = clickedSpot(view, event);
          if (forTableMenu(view, event.target, spot)) return false;
          event.preventDefault();
          view.dispatch(openMenu(view.state, event, spot));
          return true;
        },
      },
    },
  });
}
