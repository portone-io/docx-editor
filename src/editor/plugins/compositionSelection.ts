/**
 * Takes the selected text away before a composition opens over it.
 *
 * A composition normally replaces what is selected because the browser does the replacing and
 * ProseMirror reads the result back. Chrome will not touch a selection that begins at an element
 * it may not edit, though - a note's number, a preserved chip, a note reference - and it does not
 * say so: the composition lands beside the words it was meant to replace, so the selected text
 * stays standing with the syllables typed in front of it, or nothing happens at all. A footnote
 * meets it every time, since a note's first paragraph opens with the number it is drawn by, and
 * selecting the whole of a note to write it again is an ordinary thing to do.
 *
 * So the selection is deleted here rather than left to the browser, in the frame the composition
 * opens in and before ProseMirror marks the view as composing, which leaves the composition
 * starting from a caret with nothing for the browser to refuse. ProseMirror does the same for a
 * selection spanning whole blocks (`endComposition`), which is the case it already knows it cannot
 * hand over.
 *
 * The deletion is an edit like any other: the guards judge it, and a refusal leaves the browser to
 * the composition as before (`./lockedContent` ends one a refusal broke). It goes into the history
 * in the same group as the syllable that follows it, so one undo takes both.
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
