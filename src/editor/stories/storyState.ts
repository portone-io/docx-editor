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
import type { SurfaceCapabilities, SurfaceCapability } from "./storyView";

/**
 * What each of the editor's keys that puts something in needs the surface to take.
 *
 * A story is bound the key only where it takes what the key puts in, which is the very declaration
 * the toolbar's buttons are drawn from (`./storyView`), so a kind of story that takes links is
 * offered the panel both ways. The link panel is drawn over the paper, so such a kind has to say
 * where the panel stands before this key is any use to it.
 */
const CAPABILITY_KEYS: Readonly<Record<string, SurfaceCapability>> = {
  "Mod-k": "link",
  "Mod-Alt-f": "note",
};

/** The key no story of any kind is bound: a page break divides a flow a story does not have */
const KEYS_NO_STORY_TAKES: readonly string[] = ["Mod-Enter"];

function storyKeymap(takes: SurfaceCapabilities): Record<string, Command> {
  return Object.fromEntries(
    Object.entries(docxKeymap).filter(([key]) => {
      if (KEYS_NO_STORY_TAKES.includes(key)) return false;
      const needed = CAPABILITY_KEYS[key];
      return needed === undefined || takes.has(needed);
    })
  );
}

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
  /** What this kind of story takes, which the editor's own keys are bound from */
  readonly takes: SurfaceCapabilities;
}

export function storyEditorState({
  story,
  document,
  protection,
  keys,
  plugins,
  normalizers,
  takes,
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
      keymap(storyKeymap(takes)),
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
