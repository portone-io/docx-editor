/**
 * The indent commands, which work on any paragraph.
 *
 * In any paragraph but a list they move that paragraph's own left indent one step, the same as
 * Word, which is what sets them apart from `increaseListLevel` and `decreaseListLevel`.
 * In a list paragraph they move the item one level instead, exactly what Tab and Shift+Tab do.
 * A selection holding both kinds gets each paragraph treated by its own rule.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { withLeftIndent, withListNumbering } from "../../docx/paraProps";
import { toParagraphFormat } from "../../model/format";
import { editParagraphs, paragraphPPr } from "../paragraphEdits";
import { documentNumbering } from "../plugins/numberingDecorations";
import { listLevelChange, listRefOf } from "./listCommands";

/** One step of indentation: 0.5 inch, the step Word's indent buttons take */
const STEP_TWIPS = 720;

/** The left indent the paragraph is drawn at, in twips. An indent a style passes down counts too */
function leftIndentTwips(node: PMNode): number {
  const format = toParagraphFormat(node.attrs.format);
  const pt = format?.indentStartPt ?? format?.indentLeftPt ?? 0;
  return Math.round(pt * 20);
}

/** Where one step puts the paragraph's left indent. Null when it is already as far left as it goes */
function steppedLeftTwips(node: PMNode, delta: 1 | -1): number | null {
  const left = leftIndentTwips(node);
  if (delta === -1 && left <= 0) return null;
  return Math.max(0, left + delta * STEP_TWIPS);
}

function indentStep(delta: 1 | -1): Command {
  return (state, dispatch) => {
    const numbering = documentNumbering(state);
    return editParagraphs(state, dispatch, (node, styles, defaultStyleId) => {
      if (listRefOf(node)) {
        const change = listLevelChange(numbering, node, delta);
        return (
          change &&
          withListNumbering(paragraphPPr(node), change, styles, defaultStyleId)
        );
      }
      const left = steppedLeftTwips(node, delta);
      return left === null
        ? null
        : withLeftIndent(paragraphPPr(node), left, styles, defaultStyleId);
    });
  };
}

export const increaseIndent: Command = indentStep(1);

export const decreaseIndent: Command = indentStep(-1);

/** Whether anything in the selection can still move right. The toolbar disables the button when nothing can */
export function canIncreaseIndent(state: EditorState): boolean {
  return increaseIndent(state);
}

/** Whether anything in the selection can still move left. A paragraph with no indent left to give up cannot */
export function canDecreaseIndent(state: EditorState): boolean {
  return decreaseIndent(state);
}
