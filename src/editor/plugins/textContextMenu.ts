/**
 * Replaces the native text context menu and moves the selection to the clicked position. Unselected
 * table cells are left to the table context-menu plugin.
 */

import {
  type Command,
  type EditorState,
  type Plugin,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";
import { editingProtection, editsShut } from "../../schema/protectionState";
import { panelPlugin } from "./panelState";

/** Where the menu should stand (viewport coordinates) */
export interface TextMenuAnchor {
  clientX: number;
  clientY: number;
}

/** The metadata is written from wherever the transaction was built, so a point is judged as one */
function isTextMenuAnchor(value: unknown): value is TextMenuAnchor {
  if (typeof value !== "object" || value === null) return false;
  const { clientX, clientY }: Partial<TextMenuAnchor> = value;
  return typeof clientX === "number" && typeof clientY === "number";
}

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
 * merging is reached. Where the body may not be edited the table menu has nothing to offer, so
 * every click stays with this menu.
 */
function forTableMenu(
  view: EditorView,
  target: EventTarget | null,
  spot: number
): boolean {
  if (editsShut(view.state)) return false;
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
  const tr = menu.opening(state.tr, {
    clientX: event.clientX,
    clientY: event.clientY,
  });
  if (isInSelection(state, spot)) return tr;
  return tr.setSelection(TextSelection.near(tr.doc.resolve(spot)));
}

const menu = panelPlugin<TextMenuAnchor>({
  name: "docxEditorTextMenu",
  isAnchor: isTextMenuAnchor,
  // Once the document changes, what the menu was aimed at may no longer be there
  onDocChange: "close",
  props: {
    handleDOMEvents: {
      contextmenu(view, event) {
        // A reader has nothing this menu offers, so the browser menu is the better choice; a
        // commenter has the comment entry, which is why editability is not what decides
        if (editingProtection(view.state) === "readOnly") return false;
        if (!isInEditor(view, event.target)) return false;
        const spot = clickedSpot(view, event);
        // Under a shut body the entries left are the ones about the selected text: copying it
        // and commenting on it. A click landing anywhere else would open a menu with nothing to
        // do at all, so the browser's own menu is what that click is worth
        if (editsShut(view.state) && !isInSelection(view.state, spot)) {
          return false;
        }
        if (forTableMenu(view, event.target, spot)) return false;
        event.preventDefault();
        view.dispatch(openMenu(view.state, event, spot));
        return true;
      },
    },
  },
});

/** The anchor of the currently open text menu. Null when it is closed */
export function textMenuAnchor(state: EditorState): TextMenuAnchor | null {
  return menu.anchor(state);
}

/** Closes the menu. Called on running an item, Escape, clicking outside, and scrolling */
export const closeTextMenu: Command = menu.close;

export function textContextMenu(): Plugin<TextMenuAnchor | null> {
  return menu.plugin;
}
