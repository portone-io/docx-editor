/**
 * Finding a spot in a document, putting the selection there, and running a command,
 * all shared across the editing tests.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { expect } from "vitest";

/** The position just inside the first text node reading exactly this */
export function posOfText(doc: PMNode, needle: string): number {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found < 0 && node.isText && node.text === needle) found = pos + 1;
  });
  if (found < 0) throw new Error(`text not found: ${needle}`);
  return found;
}

/** A state with the caret placed, or with a range spanning the two positions */
export function select(
  state: EditorState,
  from: number,
  to = from
): EditorState {
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, from, to))
  );
}

/**
 * Actually runs the command and returns the state it produced. If it answers that it
 * has nothing to do, the test fails.
 *
 * A table command takes the same two arguments, so it goes through here as well
 */
export function runCommand(state: EditorState, command: Command): EditorState {
  let next: EditorState | null = null;
  const handled = command(state, (tr) => {
    next = state.apply(tr);
  });
  expect(handled).toBe(true);
  if (!next) throw new Error("the command produced no transaction");
  return next;
}
