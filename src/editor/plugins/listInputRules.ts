/**
 * The prefixes that start a list when they are typed at the beginning of a line.
 *
 * Word and Google Docs both turn a line beginning with a list-looking prefix into a list item
 * the moment the space after the prefix is typed, and drop the prefix. This does the same by
 * running the very commands the list buttons run, so a list started by typing is in every
 * respect the list the button makes.
 */

import { InputRule, inputRules } from "prosemirror-inputrules";
import type {
  Command,
  EditorState,
  Plugin,
  Transaction,
} from "prosemirror-state";
import { docxSchema } from "../../schema";
import {
  listRefOf,
  toggleBulletList,
  toggleNumberedList,
} from "../commands/listCommands";

/**
 * `1. ` starts a numbered list, and no other number does.
 *
 * A list started here begins counting at one, so `7. ` would come straight back as `1.` and
 * silently lose the number that was typed. Word can accept any starting number because it
 * writes that number into the list definition; until this editor does the same, only the
 * number that needs no renumbering converts and every other one stays the text it was typed as.
 */
const NUMBERED_PREFIX = /^1\. $/;

/** Both of the bullet prefixes Word and Google Docs read */
const BULLET_PREFIX = /^[-*] $/;

/** The transaction a command would dispatch. Null when the command does not apply */
function transactionOf(
  state: EditorState,
  command: Command
): Transaction | null {
  const dispatched: Transaction[] = [];
  command(state, (tr) => dispatched.push(tr));
  return dispatched[0] ?? null;
}

/**
 * Turns the paragraph the prefix was typed at the start of into a list, and takes the prefix
 * back out.
 *
 * Everything about the list is the toggle command's to decide: the numbering id it takes, that
 * the paragraph records no indentation of its own, and whether a list can be started in this
 * document at all. In a document with no numbering.xml the command applies to nothing - the
 * same answer that leaves the toolbar button disabled - and then the prefix simply stays typed.
 * The command only rewrites paragraph attributes, which leaves every position in the document
 * where it was, so the prefix is cut out of that same transaction.
 */
function startListByTyping(toggle: Command) {
  return (
    state: EditorState,
    _match: RegExpMatchArray,
    start: number,
    end: number
  ): Transaction | null => {
    const $start = state.doc.resolve(start);
    if ($start.parent.type !== docxSchema.nodes.paragraph) return null;
    // The regexp is matched against the last 500 characters of the line only, so its `^` is not
    // proof that the prefix stands at the start of the paragraph. That is asked here instead.
    if ($start.parentOffset > 0) return null;
    // An item is already given a number to look at; one that is typed into it is text
    if (listRefOf($start.parent)) return null;
    const listed = transactionOf(state, toggle);
    return listed?.delete(start, end) ?? null;
  };
}

/**
 * The plugin that watches for the prefixes.
 * A conversion is taken back by `undoInputRule`, which Backspace is bound to.
 */
export function listInputRules(): Plugin {
  return inputRules({
    rules: [
      new InputRule(NUMBERED_PREFIX, startListByTyping(toggleNumberedList)),
      new InputRule(BULLET_PREFIX, startListByTyping(toggleBulletList)),
    ],
  });
}
