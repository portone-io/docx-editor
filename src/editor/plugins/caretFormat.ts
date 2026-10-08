import type { Mark } from "prosemirror-model";
import { Plugin, TextSelection, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { docxSchema } from "../../schema";
import { formatSourceAt, sameRun } from "../formatSource";

/**
 * The marks the caret writes with where the formatting rule (`../formatSource`) parts from those
 * ProseMirror reads off the node beside it. null where the two agree.
 */
function caretMarks(tr: Transaction): readonly Mark[] | null {
  const caret =
    tr.selection instanceof TextSelection ? tr.selection.$cursor : null;
  const source = caret ? formatSourceAt(tr.doc, caret.pos) : null;
  if (!caret || source === null) return null;
  const own = caret.marks();
  const runType = docxSchema.marks.run;
  if (sameRun(runType.isInSet(own) ?? null, source.run)) return null;
  const others = runType.removeFromSet(own);
  return source.run === null ? others : source.run.addToSet(others);
}

/**
 * Gives a caret the formatting of the character the rule in `../formatSource` finds, where a mark
 * the caret was given on purpose does not stand: a caret beside something that draws nothing
 * carries on the text the rule finds, and one beside a blank takes the blank's formatting, which
 * `controlContents` then fills.
 *
 * It stands behind `controlContents`, whose stored marks for a caret left inside a control win.
 */
export function caretFormat(): Plugin {
  let live: EditorView | null = null;
  return new Plugin({
    view(view) {
      live = view;
      return {
        destroy() {
          live = null;
        },
      };
    },
    appendTransaction(transactions, _oldState, newState) {
      const moved = transactions.some((tr) => tr.docChanged || tr.selectionSet);
      // Stored marks are how ProseMirror ends a composition, so none are set under an open one
      if (!moved || newState.storedMarks !== null || live?.composing === true) {
        return null;
      }
      const tr = newState.tr;
      const marks = caretMarks(tr);
      return marks === null ? null : tr.setStoredMarks(marks);
    },
  });
}
