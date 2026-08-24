import type { Command } from "prosemirror-state";
import { docxSchema } from "../../schema";
import { replacementShut } from "../../schema/locks";

/**
 * Puts in a `w:br` of the kind its attributes name, or answers that it cannot.
 * A lock covering the selection has to answer false: the break goes in in place of whatever is
 * selected, so the guard would refuse the transaction (`schema/locks`), and a command that reports
 * it went in and did nothing draws a live control that swallows the click.
 */
function insertBreak(brAttrs: string | null): Command {
  return (state, dispatch) => {
    if (replacementShut(state.selection, state.doc)) return false;
    if (dispatch) {
      const br = docxSchema.nodes.hardBreak.create({ brAttrs });
      dispatch(state.tr.replaceSelectionWith(br).scrollIntoView());
    }
    return true;
  };
}

/** Breaks the line without splitting the paragraph */
export const insertLineBreak: Command = insertBreak(null);

/**
 * Starts the next page: the very `w:br` the editor already reads and draws in a document that
 * arrived with one.
 */
export const insertPageBreak: Command = insertBreak('w:type="page"');
