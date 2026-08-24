/**
 * Keys not handled here are taken by ProseMirror's base keymap.
 */

import { chainCommands } from "prosemirror-commands";
import { undoInputRule } from "prosemirror-inputrules";
import type { Command } from "prosemirror-state";
import { goToNextCell } from "prosemirror-tables";
import { canSplit } from "prosemirror-transform";
import { toParagraphFormat } from "../../model/format";
import { docxSchema } from "../../schema";
import { insertLineBreak, insertPageBreak } from "../commands/breakCommands";
import {
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
} from "../commands/formattingCommands";
import { redo, undo } from "../commands/historyCommands";
import { activeLinkSpan } from "../commands/linkCommands";
import {
  decreaseListLevel,
  increaseListLevel,
  leaveEmptyListItem,
} from "../commands/listCommands";
import { insertTab, moveAcrossTab } from "../commands/tabCommands";
import { openLinkPanel } from "./linkPanel";

const insertTabFromKeyboard: Command = (state, dispatch, view) => {
  // The visible link card is the next browser tab stop for a selected link.
  if (!state.selection.empty && activeLinkSpan(state)) return false;
  return insertTab(state, dispatch, view);
};

function listLevelCommand(command: Command): Command {
  return (state, dispatch, view) => {
    const format = toParagraphFormat(state.selection.$from.parent.attrs.format);
    if (!format?.numbering) return false;
    command(state, dispatch, view);
    return true;
  };
}

/**
 * Splits the paragraph on Enter.
 * The newly created paragraph inherits the original paragraph's formatting as is (the same
 * behavior as Word).
 */
const splitParagraph: Command = (state, dispatch) => {
  if (state.selection.$from.parent.type !== docxSchema.nodes.paragraph) {
    return false;
  }

  const tr = state.tr;
  if (!state.selection.empty) tr.deleteSelection();
  const at = tr.mapping.map(state.selection.$from.pos);
  const parent = tr.doc.resolve(at).parent;
  if (parent.type !== docxSchema.nodes.paragraph) return false;

  const types = [{ type: docxSchema.nodes.paragraph, attrs: parent.attrs }];
  if (!canSplit(tr.doc, at, 1, types)) return false;
  if (dispatch) dispatch(tr.split(at, 1, types).scrollIntoView());
  return true;
};

const preserveTableFollowingParagraph: Command = (state) => {
  const { $from } = state.selection;
  if (
    !state.selection.empty ||
    $from.depth !== 1 ||
    $from.parent.type !== docxSchema.nodes.paragraph ||
    $from.parent.content.size !== 0 ||
    $from.parentOffset !== 0
  ) {
    return false;
  }

  const paragraphIndex = $from.index(0);
  return (
    paragraphIndex > 0 &&
    state.doc.child(paragraphIndex - 1).type.spec.tableRole === "table"
  );
};

export const docxKeymap: Record<string, Command> = {
  Enter: chainCommands(leaveEmptyListItem, splitParagraph),
  "Shift-Enter": insertLineBreak,
  // Word and Google Docs both put a page break on this key, so it needs no learning
  "Mod-Enter": insertPageBreak,
  // The paragraph keeps adjacent tables separate without changing imported documents.
  Backspace: chainCommands(preserveTableFollowingParagraph, undoInputRule),
  "Mod-z": undo,
  "Mod-y": redo,
  "Shift-Mod-z": redo,
  "Mod-b": toggleBold,
  "Mod-i": toggleItalic,
  "Mod-u": toggleUnderline,
  // The one strikethrough key Word and Google Docs agree on
  "Mod-Shift-x": toggleStrike,
  // The link key Word and Google Docs share. It opens the panel over the selection, and with
  // nothing there to link it reports that it did nothing, so the browser keeps its own Cmd+K
  "Mod-k": openLinkPanel,
  // Inside a table, moving between cells comes first (Word does the same).
  // In a list paragraph outside a table it shifts the level; an ordinary paragraph gets a document tab.
  Tab: chainCommands(
    goToNextCell(1),
    listLevelCommand(increaseListLevel),
    insertTabFromKeyboard
  ),
  "Shift-Tab": chainCommands(
    goToNextCell(-1),
    listLevelCommand(decreaseListLevel)
  ),
  ArrowLeft: moveAcrossTab("left", false),
  ArrowRight: moveAcrossTab("right", false),
  "Shift-ArrowLeft": moveAcrossTab("left", true),
  "Shift-ArrowRight": moveAcrossTab("right", true),
};
