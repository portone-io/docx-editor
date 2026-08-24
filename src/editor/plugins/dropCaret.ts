/**
 * Where a drop lands.
 *
 * Both plugins that accept a drop from outside - plain text and image files - put what
 * they insert at the caret, so the caret has to be moved to the spot the mouse let go of
 * first.
 */

import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

/** Moves the caret to the drop position. If that position cannot be found, the caret stays where it was */
export function moveCaretToDrop(view: EditorView, event: DragEvent): void {
  const spot = view.posAtCoords({ left: event.clientX, top: event.clientY });
  if (!spot) return;
  const selection = TextSelection.near(view.state.doc.resolve(spot.pos));
  view.dispatch(view.state.tr.setSelection(selection));
}
