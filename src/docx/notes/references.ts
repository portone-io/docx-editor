/**
 * Where the body refers to its notes.
 *
 * Three readers ask: the notes writers collect the ids each part is referred to by, the numbering
 * counts the references section by section (`./numbering`), and the export invariants look up the
 * first reference to one note, which is where a problem about a note's own text is reported
 * (`docx/invariants`). One walk here is what keeps the three reading the same nodes in the same
 * order.
 */

import type { Node as PMNode } from "prosemirror-model";
import { NOTE_KINDS, type NoteKey, storyKey } from "../../schema/stories";

/** One reference the body makes to a note */
export interface PlacedNoteReference {
  readonly node: PMNode;
  /** Where the reference stands in the document */
  readonly pos: number;
  /** The top-level block it stands in, by index, which is what a section covers a run of */
  readonly block: number;
  /** The note it names, and null where it names no kind of note or carries no id */
  readonly key: NoteKey | null;
}

function keyOf(node: PMNode): NoteKey | null {
  const kind = NOTE_KINDS.find((candidate) => candidate === node.attrs.kind);
  const id: unknown = node.attrs.id;
  return kind === undefined || typeof id !== "string"
    ? null
    : storyKey(kind, id);
}

/**
 * Every `noteReference` of the body in document order, until `visit` answers false.
 *
 * Only the document node is walked: WordprocessingML puts a reference in the body and a note's own
 * text in a part of its own, so a side story holds none.
 */
export function eachNoteReference(
  doc: PMNode,
  visit: (reference: PlacedNoteReference) => boolean | void
): void {
  let running = true;
  doc.forEach((block, offset, index) => {
    if (!running) return;
    block.descendants((node, pos) => {
      if (!running) return false;
      if (node.type.name !== "noteReference") return true;
      running =
        visit({
          node,
          pos: offset + 1 + pos,
          block: index,
          key: keyOf(node),
        }) !== false;
      return false;
    });
  });
}

/**
 * Where the body first refers to each of these notes, and nothing for one it refers to nowhere.
 * The walk stops as soon as every note asked about has been placed.
 */
export function firstNoteReferences(
  doc: PMNode,
  wanted: ReadonlySet<NoteKey>
): ReadonlyMap<NoteKey, number> {
  const first = new Map<NoteKey, number>();
  if (wanted.size === 0) return first;
  eachNoteReference(doc, ({ key, pos }) => {
    if (key !== null && wanted.has(key) && !first.has(key)) first.set(key, pos);
    return first.size < wanted.size;
  });
  return first;
}
