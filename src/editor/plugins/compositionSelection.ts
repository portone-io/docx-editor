/**
 * Takes the selected text away before a composition opens over it.
 *
 * A composition normally replaces what is selected because the browser does the replacing and
 * ProseMirror reads the result back. Chrome will not touch a selection that begins at an element
 * it may not edit - the number a note is drawn by, a note reference - and does not say so: the
 * composition lands beside the words it was meant to replace, or nothing happens at all. A note
 * meets it every time, since its first paragraph opens with its number and selecting the whole of
 * a note to write it again is an ordinary thing to do.
 *
 * So the selection is deleted here, in the frame the composition opens in and before ProseMirror
 * marks the view as composing, which leaves the composition starting from a caret with nothing
 * for the browser to refuse. ProseMirror does the same for a selection spanning whole blocks
 * (`endComposition`).
 *
 * The deletion is an edit like any other, so the guards judge it: across a bookmark marker it is
 * refused, and the composition is left to the browser where it stood.
 */

import { deleteSelection } from "prosemirror-commands";
import { Plugin } from "prosemirror-state";

export function compositionSelection(): Plugin {
  return new Plugin({
    props: {
      handleDOMEvents: {
        compositionstart(view) {
          if (view.state.selection.empty) return false;
          deleteSelection(view.state, (tr) => view.dispatch(tr));
          // The composition is still the browser's to open
          return false;
        },
      },
    },
  });
}
