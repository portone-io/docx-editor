/**
 * The state one side story is edited in.
 *
 * A story is a document of the same schema as the body, so it is edited by the same plugins: the
 * guards, the clipboard, the tabs, the display derivation. What it is not is a document of its
 * own, so the ones that answer for a whole document stay out - the history, which the main
 * document keeps for both (`./storyView`), the note and comment plugins, whose subjects a story
 * does not hold, and the page decorations, since a story is drawn wherever its host puts it.
 *
 * The snapshot handed in is the main document's, which is what makes a run resolve here as it
 * would in the body. Its caller narrows the two values a story answers differently: a story has
 * nowhere to write a list definition, and the paper it wraps at is the one its anchor stands on.
 */

import { baseKeymap } from "prosemirror-commands";
import { dropCursor } from "prosemirror-dropcursor";
import { keymap } from "prosemirror-keymap";
import type { Node as PMNode } from "prosemirror-model";
import { type Command, EditorState, type Plugin } from "prosemirror-state";
import { tableEditing } from "prosemirror-tables";
import type { EditingProtection } from "../../schema/protection";
import type { SliceNormalizer } from "../clipboard/normalizers";
import { docxClipboard } from "../clipboard/plugin";
import { type EditorDocument, editorDocument } from "../editorDocument";
import {
  displayDerivation,
  withDerivedDisplay,
} from "../plugins/displayDerivation";
import { documentProtection } from "../plugins/documentProtection";
import { docxKeymap } from "../plugins/keymap";
import { lockedContent } from "../plugins/lockedContent";
import { numberingMarkers } from "../plugins/numberingDecorations";
import { tabCaret } from "../plugins/tabCaret";
import { tabDecorations } from "../plugins/tabDecorations";
import { tabLayout } from "../plugins/tabLayout";
import { tabPointer } from "../plugins/tabPointer";

/**
 * The editor's keys that mean nothing inside a story: a page break, which only a body starts, the
 * link panel, which floats over the paper and would have nowhere to stand, and a note, which is
 * called from the document story alone.
 */
const KEYS_OUTSIDE_A_STORY: readonly string[] = [
  "Mod-Enter",
  "Mod-k",
  "Mod-Alt-f",
];

const STORY_KEYMAP: Record<string, Command> = Object.fromEntries(
  Object.entries(docxKeymap).filter(
    ([key]) => !KEYS_OUTSIDE_A_STORY.includes(key)
  )
);

export interface StoryStateOptions {
  /** The story as the main document holds it */
  readonly story: PMNode;
  /** The main document's snapshot, which the story's display values are resolved against */
  readonly document: EditorDocument;
  readonly protection: EditingProtection;
  /** The keys the surface binds ahead of the editor's own: undo, redo and Escape */
  readonly keys: Record<string, Command>;
  /** What the kind of story adds, which is asked before the editor's own keymaps */
  readonly plugins: readonly Plugin[];
  readonly normalizers: readonly SliceNormalizer[];
}

export function storyEditorState({
  story,
  document,
  protection,
  keys,
  plugins,
  normalizers,
}: StoryStateOptions): EditorState {
  return EditorState.create({
    doc: withDerivedDisplay(story, document),
    plugins: [
      editorDocument(document),
      lockedContent(),
      documentProtection({ protection, author: null, editableComments: "own" }),
      // A kind's own key rule stands ahead of the keymaps below: the base keymap answers
      // Backspace and Enter for every document, so a rule about an empty story would never be
      // asked behind it
      ...plugins,
      keymap(keys),
      keymap(STORY_KEYMAP),
      keymap(baseKeymap),
      dropCursor(),
      docxClipboard({ normalizers }),
      tableEditing(),
      tabDecorations(),
      tabPointer(),
      tabLayout(),
      tabCaret(),
      displayDerivation(),
      numberingMarkers(),
    ],
  });
}
