/** Document-level readers for imported footnotes and endnotes. */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { type StoryKey, storyKey, storyOf, storyText } from "../../docx/story";
import type { NoteKind } from "../../schema/stories";
import { documentProjection } from "../plugins/documentProjection";

export interface DocumentNote {
  readonly kind: NoteKind;
  readonly id: string;
  readonly label: string;
  readonly text: string;
  readonly referencePos: number;
}

/** One note the body refers to, as the notes drawn around the page draw it */
export interface NoteRow {
  readonly key: StoryKey;
  readonly kind: NoteKind;
  readonly label: string;
  /** The note's story, the same node for as long as it says the same thing (`schema/stories`) */
  readonly story: PMNode;
  /** Where the first reference to it stands, which is the place in the text it is called from */
  readonly referencePos: number;
}

function stringAttr(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function noteKindOf(node: PMNode): NoteKind {
  return node.attrs.kind === "endnote" ? "endnote" : "footnote";
}

interface NoteProjection {
  notes: readonly DocumentNote[];
  /**
   * The body of each note on the marker pointing at it, which is the tooltip a reader gets without
   * looking away from the line. It is drawn rather than written on the node, because what a note
   * says is a story on the document and not an attr of the reference.
   */
  tooltips: DecorationSet;
  /** The footnotes whose story the document holds, by story key, in first-reference order */
  footnotes: ReadonlyMap<StoryKey, NoteRow>;
  /** The endnotes whose story the document holds, in first-reference order */
  endnotes: readonly NoteRow[];
}

function deriveNotes(doc: PMNode): NoteProjection {
  const notes: DocumentNote[] = [];
  const tooltips: Decoration[] = [];
  const footnotes = new Map<StoryKey, NoteRow>();
  const endnotes: NoteRow[] = [];
  const seen = new Set<string>();
  doc.descendants((node, pos) => {
    if (node.type.name !== "noteReference") return true;
    const id = stringAttr(node.attrs.id);
    if (id === null) return true;
    const kind = noteKindOf(node);
    const key = storyKey(kind, id);
    const story = storyOf(doc, key);
    const text = storyText(story);
    if (text !== "") {
      tooltips.push(Decoration.node(pos, pos + node.nodeSize, { title: text }));
    }
    if (seen.has(key)) return true;
    seen.add(key);
    const label = stringAttr(node.attrs.label) ?? "?";
    notes.push({ kind, id, label, text, referencePos: pos });
    if (story !== null) {
      const row: NoteRow = { key, kind, label, story, referencePos: pos };
      if (kind === "footnote") footnotes.set(key, row);
      else endnotes.push(row);
    }
    return true;
  });
  return {
    notes,
    tooltips: DecorationSet.create(doc, tooltips),
    footnotes,
    endnotes,
  };
}

/**
 * Which notes the document refers to is the document's to decide, so the list is worked out once
 * per edit (`editor/plugins/documentProjection`) rather than once per render of the notes drawn
 * around the page.
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
