import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { toParagraphFormat } from "../../model/format";
import { docxSchema } from "../../schema";
import { replacementShut } from "../../schema/locks";

function isOrdinaryParagraph(state: EditorState): boolean {
  const { $from } = state.selection;
  if (
    $from.parent.type !== docxSchema.nodes.paragraph ||
    toParagraphFormat($from.parent.attrs.format)?.numbering
  ) {
    return false;
  }
  for (let depth = $from.depth - 1; depth > 0; depth -= 1) {
    if ($from.node(depth).type.spec.tableRole === "cell") return false;
  }
  return true;
}

/** Inserts an editable OOXML tab in an ordinary paragraph. */
export const insertTab: Command = (state, dispatch) => {
  const { selection } = state;
  if (
    !(selection instanceof TextSelection) ||
    !selection.$from.sameParent(selection.$to) ||
    !isOrdinaryParagraph(state) ||
    replacementShut(selection, state.doc)
  ) {
    return false;
  }
  if (dispatch) {
    const inherited =
      state.storedMarks ??
      (selection.empty
        ? selection.$from.marks()
        : (selection.$from.marksAcross(selection.$to) ?? []));
    const marks = inherited.filter(
      (mark) => mark.type !== docxSchema.marks.tab
    );
    const tab = docxSchema.text(
      "\t",
      docxSchema.marks.tab.create().addToSet(marks)
    );
    dispatch(state.tr.replaceSelectionWith(tab, false).scrollIntoView());
  }
  return true;
};

/** Moves one logical text position only when that position contains a tab. */
export function moveAcrossTab(
  visualDirection: "left" | "right",
  extend: boolean
): Command {
  return (state, dispatch) => {
    const { selection } = state;
    if (
      !(selection instanceof TextSelection) ||
      (!extend && !selection.empty)
    ) {
      return false;
    }
    const { $head } = selection;
    if ($head.parent.type !== docxSchema.nodes.paragraph) return false;
    const rtl =
      toParagraphFormat($head.parent.attrs.format)?.direction === "rtl";
    const forward = visualDirection === "right" ? !rtl : rtl;
    const delta = forward ? 1 : -1;
    const nextOffset = $head.parentOffset + delta;
    if (nextOffset < 0 || nextOffset > $head.parent.content.size) return false;
    const from = delta > 0 ? $head.pos : $head.pos - 1;
    if (state.doc.textBetween(from, from + 1, "", "") !== "\t") return false;
    if (dispatch) {
      const anchor = extend ? selection.anchor : $head.pos + delta;
      dispatch(
        state.tr
          .setSelection(
            TextSelection.create(state.doc, anchor, $head.pos + delta)
          )
          .scrollIntoView()
      );
    }
    return true;
  };
}
