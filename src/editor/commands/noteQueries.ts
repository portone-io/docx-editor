/** Document-level readers for imported footnotes and endnotes. */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import type { NoteKind } from "../../docx/notes";
import { documentProjection } from "../plugins/documentProjection";

export interface DocumentNote {
  kind: NoteKind;
  id: string;
  label: string;
  text: string;
  referencePos: number;
}

function stringAttr(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function notesIn(doc: PMNode): readonly DocumentNote[] {
  const notes: DocumentNote[] = [];
  const seen = new Set<string>();
  doc.descendants((node, pos) => {
    if (node.type.name !== "noteReference") return true;
    const id = stringAttr(node.attrs.id);
    if (id === null) return true;
    const kind: NoteKind =
      node.attrs.kind === "endnote" ? "endnote" : "footnote";
    const key = `${kind}:${id}`;
    if (seen.has(key)) return true;
    seen.add(key);
    notes.push({
      kind,
      id,
      label: stringAttr(node.attrs.label) ?? "?",
      text: stringAttr(node.attrs.text) ?? "",
      referencePos: pos,
    });
    return true;
  });
  return notes;
}

/**
 * Which notes the document refers to is the document's to decide, so the list is worked out once
 * per edit (`editor/plugins/documentProjection`) rather than once per render of the panel under
 * the page.
 */
export const noteProjection = documentProjection<readonly DocumentNote[]>(
  "docxEditorNotes",
  notesIn
);

/** The distinct notes referenced by the main document story, in first-reference order. */
export function documentNotes(state: EditorState): readonly DocumentNote[] {
  return noteProjection.read(state);
}
