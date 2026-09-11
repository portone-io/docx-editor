/** Document-level readers for imported footnotes and endnotes. */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import type { NoteKind } from "../../docx/notes/reading";
import { storyKey, storyOf, storyText } from "../../docx/story";
import { documentProjection } from "../plugins/documentProjection";

export interface DocumentNote {
  readonly kind: NoteKind;
  readonly id: string;
  readonly label: string;
  readonly text: string;
  readonly referencePos: number;
}

function stringAttr(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function noteKindOf(node: PMNode): NoteKind {
  return node.attrs.kind === "endnote" ? "endnote" : "footnote";
}

/** What one note says, read off the story the document holds it in (`docx/story`) */
function noteText(doc: PMNode, kind: NoteKind, id: string): string {
  return storyText(storyOf(doc, storyKey(kind, id)));
}

interface NoteProjection {
  notes: readonly DocumentNote[];
  /**
   * The body of each note on the marker pointing at it, which is the tooltip a reader gets without
   * looking away from the line. It is drawn rather than written on the node, because what a note
   * says is a story on the document and not an attr of the reference.
   */
  tooltips: DecorationSet;
}

function deriveNotes(doc: PMNode): NoteProjection {
  const notes: DocumentNote[] = [];
  const tooltips: Decoration[] = [];
  const seen = new Set<string>();
  doc.descendants((node, pos) => {
    if (node.type.name !== "noteReference") return true;
    const id = stringAttr(node.attrs.id);
    if (id === null) return true;
    const kind = noteKindOf(node);
    const text = noteText(doc, kind, id);
    if (text !== "") {
      tooltips.push(Decoration.node(pos, pos + node.nodeSize, { title: text }));
    }
    const key = storyKey(kind, id);
    if (seen.has(key)) return true;
    seen.add(key);
    notes.push({
      kind,
      id,
      label: stringAttr(node.attrs.label) ?? "?",
      text,
      referencePos: pos,
    });
    return true;
  });
  return { notes, tooltips: DecorationSet.create(doc, tooltips) };
}

/**
 * Which notes the document refers to is the document's to decide, so the list is worked out once
 * per edit (`editor/plugins/documentProjection`) rather than once per render of the panel under
 * the page.
 */
export const noteProjection = documentProjection<NoteProjection>(
  "docxEditorNotes",
  deriveNotes,
  {
    // The set the last edit derived. Written as a method so that `this` is the plugin holding it
    decorations(state) {
      return this.getState(state)?.tooltips;
    },
  }
);

/** The distinct notes referenced by the main document story, in first-reference order. */
export function documentNotes(state: EditorState): readonly DocumentNote[] {
  return noteProjection.read(state).notes;
}
