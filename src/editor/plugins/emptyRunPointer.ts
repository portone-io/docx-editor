import { Plugin, TextSelection } from "prosemirror-state";
import { isEmptyRun } from "../../schema/emptyRuns";

/**
 * Puts the caret after a run holding no characters when the blank it is drawn as is pressed.
 *
 * Text takes its formatting from the nearest character shown before the caret (`../formatSource`),
 * so a caret dropped on the left half of the blank would carry on the words before it rather than
 * fill the blank. After it, what is written there takes the blank's formatting and fills it.
 */
export function emptyRunPointer(): Plugin {
  return new Plugin({
    props: {
      handleClickOn(view, _pos, node, nodePos, _event, direct) {
        if (!direct || !isEmptyRun(node)) return false;
        const { state } = view;
        view.dispatch(
          state.tr.setSelection(
            TextSelection.create(state.doc, nodePos + node.nodeSize)
          )
        );
        return true;
      },
    },
  });
}
