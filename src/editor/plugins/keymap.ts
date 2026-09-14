/**
 * Keys not handled here are taken by ProseMirror's base keymap.
 */

import { chainCommands } from "prosemirror-commands";
import { undoInputRule } from "prosemirror-inputrules";
import { keydownHandler } from "prosemirror-keymap";
import type { Attrs, Node as PMNode } from "prosemirror-model";
import { type Command, Plugin } from "prosemirror-state";
import { goToNextCell } from "prosemirror-tables";
import { canSplit } from "prosemirror-transform";
import { splitParagraphAttrs } from "../../docx/cloning";
import { toParagraphFormat } from "../../model/format";
import { docxSchema } from "../../schema";
import { insertLineBreak, insertPageBreak } from "../commands/breakCommands";
import { insertFootnote } from "../commands/footnoteCommands";
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
 *
 * The newly created paragraph inherits the original paragraph's formatting as is (the same
 * behavior as Word), but not what named the original: `docx/cloning` decides which half carries
 * the identifiers and which one carries the section break.
 *
 * The half left standing is rewritten only where the policy actually changed something, so an
 * ordinary Enter stays the single step it was and reaches no further than the spot it split.
 */
const splitParagraph: Command = (state, dispatch) => {
  if (state.selection.$from.parent.type !== docxSchema.nodes.paragraph) {
    return false;
  }

  const tr = state.tr;
  if (!state.selection.empty) tr.deleteSelection();
  const at = tr.mapping.map(state.selection.$from.pos);
  const $at = tr.doc.resolve(at);
  const parent = $at.parent;
  if (parent.type !== docxSchema.nodes.paragraph) return false;

  const { before, after } = splitParagraphAttrs(parent);
  const types = [{ type: docxSchema.nodes.paragraph, attrs: after }];
  if (!canSplit(tr.doc, at, 1, types)) return false;
  if (dispatch) {
    const beforePos = $at.before();
    tr.split(at, 1, types);
    if (!keptEveryAttr(parent, before)) {
      tr.setNodeMarkup(beforePos, null, before);
    }
    dispatch(tr.scrollIntoView());
  }
  return true;
};

/**
 * Whether the policy handed back every attr of the paragraph exactly as it stood.
 *
 * A rewritten value is a fresh string or null where the original held something else, and a value
 * carried over is the very one the node holds, so identity answers this without walking into the
 * derived display values.
 */
function keptEveryAttr(paragraph: PMNode, attrs: Attrs): boolean {
  return Object.keys(paragraph.attrs).every(
    (name) => attrs[name] === paragraph.attrs[name]
  );
}

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

/** The undo and redo keys on their own, so that a view which took editing away can still run them */
const historyKeymap: Record<string, Command> = {
  "Mod-z": undo,
  "Mod-y": redo,
  "Shift-Mod-z": redo,
};

/**
 * Keeps undo and redo on their keys where the keymap above cannot reach them.
 *
 * A view that took editing away drops every keydown before the keymap plugin sees it
 * (`view.editable || !(event.type in editHandlers)` in prosemirror-view), so a commenter, who is
 * given no caret, would have no way to take a comment back. A DOM handler is asked ahead of that
 * check. An editable view leaves the keys to the keymap plugin, which would otherwise run them a
 * second time.
 */
export function historyKeys(): Plugin {
  const runHistoryKeys = keydownHandler(historyKeymap);
  return new Plugin({
    props: {
      handleDOMEvents: {
        keydown: (view, event) => {
          if (view.editable || !runHistoryKeys(view, event)) return false;
          event.preventDefault();
          return true;
        },
      },
    },
  });
}

export const docxKeymap: Record<string, Command> = {
  Enter: chainCommands(leaveEmptyListItem, splitParagraph),
  "Shift-Enter": insertLineBreak,
  // Word and Google Docs both put a page break on this key, so it needs no learning
  "Mod-Enter": insertPageBreak,
  // The paragraph keeps adjacent tables separate without changing imported documents.
  Backspace: chainCommands(preserveTableFollowingParagraph, undoInputRule),
  ...historyKeymap,
  "Mod-b": toggleBold,
  "Mod-i": toggleItalic,
  "Mod-u": toggleUnderline,
  // The one strikethrough key Word and Google Docs agree on
  "Mod-Shift-x": toggleStrike,
  // The link key Word and Google Docs share. It opens the panel over the selection, and with
  // nothing there to link it reports that it did nothing, so the browser keeps its own Cmd+K
  "Mod-k": openLinkPanel,
  // Google Docs' footnote key, which puts the caret inside the footnote it adds
  "Mod-Alt-f": insertFootnote,
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
