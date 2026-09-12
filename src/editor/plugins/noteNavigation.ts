/**
 * Getting into a note and asking for one to be opened.
 *
 * A press on a note's number or a command asks for a note, and neither can mount anything itself,
 * since what draws the notes is the surface around the editor; so the request is held in the state
 * and the surface reads it back (`requestedNote`). The nonce is what makes a second request for
 * the note already open reach the surface at all.
 *
 * The request moves the caret to just after the reference as it goes in, which is where Escape
 * hands it back to (`editor/notes/noteSurface`), so leaving a note needs nothing remembered.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  Plugin,
  PluginKey,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { docxSchema } from "../../schema";
import {
  EDITABLE_NOTE_KINDS,
  type NoteKind,
  type StoryKey,
  storyKey,
  storyNodeOf,
} from "../../schema/stories";
import { documentOf } from "../editorDocument";

/** Which note the reader asked for, and which asking it was */
export interface RequestedNote {
  readonly key: StoryKey;
  /** Rises with every request, so asking twice for one note is two requests */
  readonly nonce: number;
}

const navigationKey = new PluginKey<RequestedNote | null>(
  "docxEditorNoteNavigation"
);

/** One reference to a note, as the document holds it */
export interface PlacedNoteReference {
  readonly pos: number;
  readonly node: PMNode;
}

/** Where the first reference to this note stands, and null where the document holds none */
export function noteReferenceAt(
  doc: PMNode,
  kind: NoteKind,
  id: string
): PlacedNoteReference | null {
  let found: PlacedNoteReference | null = null;
  doc.descendants((node, pos) => {
    if (found !== null) return false;
    if (
      node.type === docxSchema.nodes.noteReference &&
      node.attrs.kind === kind &&
      node.attrs.id === id
    ) {
      found = { pos, node };
    }
    return found === null;
  });
  return found;
}

function isRequest(value: unknown): value is RequestedNote {
  if (typeof value !== "object" || value === null) return false;
  return (
    typeof Reflect.get(value, "key") === "string" &&
    typeof Reflect.get(value, "nonce") === "number"
  );
}

let asked = 0;

/** The transaction with the note it opens recorded on it, which `insertFootnote` also carries */
export function requestNote<T extends Transaction>(tr: T, key: StoryKey): T {
  asked += 1;
  return tr.setMeta(navigationKey, { key, nonce: asked });
}

/** The note last asked for, and null where none was. Null for a state built without the plugin */
export function requestedNote(state: EditorState): RequestedNote | null {
  return navigationKey.getState(state) ?? null;
}

/**
 * Puts the caret just after the note's reference and asks for the note to be opened. It applies
 * under every mode, since opening a note is reading it; what a reader may do once it stands open
 * is the story view's question (`editor/stories`).
 */
export function openNoteCommand(kind: NoteKind, id: string): Command {
  return (state, dispatch) => {
    const key = storyKey(kind, id);
    if (documentOf(state).specialNotes.has(key)) return false;
    if (storyNodeOf(state.doc, key) === null) return false;
    const found = noteReferenceAt(state.doc, kind, id);
    if (found === null) return false;
    if (dispatch) {
      const after = found.pos + found.node.nodeSize;
      const tr = state.tr.setSelection(
        TextSelection.near(state.doc.resolve(after))
      );
      dispatch(requestNote(tr, key).scrollIntoView());
    }
    return true;
  };
}

/**
 * Holds the note last asked for, and turns a press on a note's number into that asking.
 *
 * One press is what opens a note, which is what Google Docs does: the number is a small
 * superscript, and a press on one is far more often a reader going to the note than a caret being
 * placed on the digit. A reference to a kind no writer puts back opens nothing.
 */
export function noteNavigation(
  kinds: readonly NoteKind[] = EDITABLE_NOTE_KINDS
): Plugin<RequestedNote | null> {
  return new Plugin<RequestedNote | null>({
    key: navigationKey,
    state: {
      init: () => null,
      apply: (tr, held) => {
        const next: unknown = tr.getMeta(navigationKey);
        return isRequest(next) ? next : held;
      },
    },
    props: {
      handleClickOn(view, _pos, node, _nodePos, _event, direct) {
        if (!direct || node.type !== docxSchema.nodes.noteReference) {
          return false;
        }
        const kind = kinds.find((named) => named === node.attrs.kind);
        const id: unknown = node.attrs.id;
        if (kind === undefined || typeof id !== "string") return false;
        return openNoteCommand(kind, id)(view.state, view.dispatch);
      },
    },
  });
}
