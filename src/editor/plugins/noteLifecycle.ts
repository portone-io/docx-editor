/**
 * Keeps each note of an editable kind in step with the references that call it.
 *
 * A note is written once and called from the text by its id, so what an edit does to the text
 * decides what becomes of the note. An edit that takes away the last reference to a note takes the
 * note away with it. An edit that copies a reference gives the copy a new id and a copy of the
 * note, since two references naming one id are one note drawn with one number where the reader
 * made two. A reference moved whole keeps its id and its note. Either change goes in a transaction
 * appended to the edit, so it lands in the edit's history event and one undo takes both back.
 *
 * Only what the edit did is read. An entry the file arrived with and no reference to is never
 * reached, a reference naming no note is left naming none, and the separators a notes part lays
 * the page out with are left as they are. The notes part writer relies on the first of these: a
 * note disappears only together with its last reference (`docx/storyParts`).
 */

import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, Plugin, type Transaction } from "prosemirror-state";
import { DocAttrStep, Mapping } from "prosemirror-transform";
import {
  copiedNoteStory,
  nextNoteId,
  takenNoteIds,
} from "../../docx/notes/newNote";
import { setStory, withoutStories } from "../../docx/story";
import { displayOnly } from "../../schema/displayDerivation";
import { stepReaches } from "../../schema/guards";
import {
  EDITABLE_NOTE_KINDS,
  type NoteKind,
  type StoryKey,
  storyKey,
  storyNodeOf,
} from "../../schema/stories";
import { documentOf } from "../editorDocument";

interface PlacedReference {
  readonly kind: NoteKind;
  readonly key: StoryKey;
  readonly pos: number;
}

function kindOf(node: PMNode, kinds: readonly NoteKind[]): NoteKind | null {
  if (node.type.name !== "noteReference") return null;
  return kinds.find((kind) => kind === node.attrs.kind) ?? null;
}

/** Every reference to a note of these kinds, in document order */
function referencesIn(
  doc: PMNode,
  kinds: readonly NoteKind[]
): PlacedReference[] {
  const found: PlacedReference[] = [];
  doc.descendants((node, pos) => {
    const kind = kindOf(node, kinds);
    const id: unknown = node.attrs.id;
    if (kind !== null && typeof id === "string") {
      found.push({ kind, key: storyKey(kind, id), pos });
    }
    return true;
  });
  return found;
}

function countsOf(
  references: readonly PlacedReference[]
): ReadonlyMap<StoryKey, number> {
  return references.reduce(
    (counts, { key }) => counts.set(key, (counts.get(key) ?? 0) + 1),
    new Map<StoryKey, number>()
  );
}

/**
 * Whether the transaction puts in, takes out or rewrites such a reference.
 *
 * A document attr step writes a story or a section and moves no reference, although `stepReaches`
 * answers a step it does not know as reaching one; typing into a note is such a step. A
 * re-derivation of display values changes no id and moves nothing.
 */
function reachesReferences(
  tr: Transaction,
  kinds: readonly NoteKind[]
): boolean {
  if (!tr.docChanged || tr.getMeta(displayOnly) === true) return false;
  return tr.steps.some(
    (step, index) =>
      !(step instanceof DocAttrStep) &&
      stepReaches(
        step,
        tr.docs[index],
        tr.docs[index + 1] ?? tr.doc,
        (node) => kindOf(node, kinds) !== null
      )
  );
}

/**
 * The references the edit copied: for every note it left more references to than it found, those
 * past the count it found.
 *
 * The references the edit carried from where they stood are counted first, so the one a copy was
 * made from keeps its id wherever the copy lands; a reference the edit rebuilt rather than carried
 * counts after them, in document order.
 */
function copiedReferences(
  transactions: readonly Transaction[],
  before: readonly PlacedReference[],
  after: readonly PlacedReference[]
): PlacedReference[] {
  const found = countsOf(before);
  const left = countsOf(after);
  const mapping = new Mapping(transactions.flatMap((tr) => tr.mapping.maps));
  const carried = new Set(
    before.flatMap(({ key, pos }) => {
      const mapped = mapping.mapResult(pos, 1);
      return mapped.deletedAfter ? [] : [`${key}@${mapped.pos}`];
    })
  );
  const wasCarried = ({ key, pos }: PlacedReference) =>
    carried.has(`${key}@${pos}`);
  return [...left].flatMap(([key, count]) => {
    const kept = found.get(key) ?? 0;
    if (kept === 0 || count <= kept) return [];
    const standing = after.filter((reference) => reference.key === key);
    return [
      ...standing.filter(wasCarried),
      ...standing.filter((reference) => !wasCarried(reference)),
    ].slice(kept);
  });
}

function settled(
  transactions: readonly Transaction[],
  oldState: EditorState,
  newState: EditorState,
  kinds: readonly NoteKind[]
): Transaction | null {
  const document = documentOf(newState);
  const isNote = (key: StoryKey) =>
    !document.specialNotes.has(key) && storyNodeOf(newState.doc, key) !== null;
  const before = referencesIn(oldState.doc, kinds);
  const after = referencesIn(newState.doc, kinds);
  const left = countsOf(after);
  const gone = [...countsOf(before).keys()].filter(
    (key) => !left.has(key) && isNote(key)
  );
  const copies = copiedReferences(transactions, before, after)
    .filter(({ key }) => isNote(key))
    .sort((a, b) => a.pos - b.pos);
  if (gone.length === 0 && copies.length === 0) return null;

  const tr = withoutStories(newState.tr, gone);
  const taken = new Map<NoteKind, Set<string>>();
  for (const copy of copies) {
    const story = storyNodeOf(newState.doc, copy.key);
    if (story === null) continue;
    const ids =
      taken.get(copy.kind) ??
      takenNoteIds(newState.doc, copy.kind, document.reservedNoteKeys);
    taken.set(copy.kind, ids);
    const id = nextNoteId(ids);
    ids.add(id);
    // The XML the reference arrived as names the id it was copied with, so the copy is written
    // from its attrs instead
    tr.setNodeAttribute(copy.pos, "id", id).setNodeAttribute(
      copy.pos,
      "referenceXml",
      null
    );
    setStory(
      tr,
      storyKey(copy.kind, id),
      copiedNoteStory(story, document.session)
    );
  }
  return tr;
}

export function noteLifecycle(
  kinds: readonly NoteKind[] = EDITABLE_NOTE_KINDS
): Plugin {
  return new Plugin({
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((tr) => reachesReferences(tr, kinds))) {
        return null;
      }
      return settled(transactions, oldState, newState, kinds);
    },
  });
}
