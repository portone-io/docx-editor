/**
 * Where the protection (`./protection`) stands in an editor state, and the questions commands ask
 * of it. Kept apart from the judgements so that those stay free of editor state for the `./core`
 * entry.
 */

import { type EditorState, PluginKey } from "prosemirror-state";
import type { EditingProtection, ProtectionState } from "./protection";

/** What a state built without the protection plugin runs under, which is what every consumer had before it existed */
export const UNPROTECTED: ProtectionState = {
  protection: "none",
  authorId: null,
  editableComments: "own",
};

export const protectionKey = new PluginKey<ProtectionState>(
  "docxEditorProtection"
);

export function protectionOf(state: EditorState): ProtectionState {
  return protectionKey.getState(state) ?? UNPROTECTED;
}

export function editingProtection(state: EditorState): EditingProtection {
  return protectionOf(state).protection;
}

/**
 * Whether the protection shuts every edit to the body, comments aside.
 *
 * This is the question a command that changes text, formatting or structure asks before it reports
 * that it applies, the way it asks the lock predicates (`./locks`): a command reporting true and
 * then being refused by the guard draws a live control that swallows the click.
 */
export function editsShut(state: EditorState): boolean {
  const protection = editingProtection(state);
  switch (protection) {
    case "none":
      return false;
    case "readOnly":
    case "comments":
      return true;
    default: {
      const unmodelled: never = protection;
      return unmodelled;
    }
  }
}
