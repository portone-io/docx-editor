/**
 * What a footnote or an endnote draws and does differently from any other story.
 *
 * A note's story opens with a mark of its own (`w:footnoteRef`, `w:endnoteRef`), which the file
 * keeps as a preserved chip. Word draws that mark as the number the reference carries, so it is
 * drawn here as that label in superscript, the way the reference in the body is.
 *
 * The rest is what the story view asks a kind of story for (`editor/stories`): the host every
 * edit leaves through, and the extensions that give a note its chip, its key rule, the number it
 * keeps through an edit, and what a paste into one may carry.
 */

import { keymap } from "prosemirror-keymap";
import {
  type DOMOutputSpec,
  DOMSerializer,
  Mark,
  type Node as PMNode,
} from "prosemirror-model";
import { type Command, Plugin, TextSelection } from "prosemirror-state";
import type { EditorView, NodeViewConstructor } from "prosemirror-view";
import { NOTE_NUMBER_ELEMENTS } from "../../docx/notes/newNote";
import { docxSchema } from "../../schema";
import { editShut, transactionAllowed } from "../../schema/guards";
import { editsShut } from "../../schema/protectionState";
import { noteKeyOf, type StoryKey } from "../../schema/stories";
import { editorClassNames } from "../../styles/classNames";
import {
  detachAnchors,
  dropSourceIdentity,
  mapSliceNodes,
  rederiveDisplay,
  rekeyNumbering,
  type SliceNormalizer,
} from "../clipboard/normalizers";
import { redo, undo } from "../commands/historyCommands";
import { noteBodyCommand } from "../commands/noteCommands";
import { noteReferenceAt } from "../plugins/noteNavigation";
import type { StoryNodeSpecs } from "../stories/storyMarkup";
import {
  NO_CAPABILITY,
  type StoryExtensions,
  type StoryHost,
} from "../stories/storyView";

/** Whether this is the mark a note's entry opens with, which Word draws as the note's number */
function isOwnMark(node: PMNode): boolean {
  const element: unknown = node.attrs.element;
  return (
    node.type === docxSchema.nodes.rawRunContent &&
    typeof element === "string" &&
    NOTE_NUMBER_ELEMENTS.includes(element)
  );
}

function drawnBySchema(node: PMNode): DOMOutputSpec {
  return docxSchema.nodes.rawRunContent.spec.toDOM?.(node) ?? ["span"];
}

function markSpec(node: PMNode, labelOf: () => string): DOMOutputSpec {
  return isOwnMark(node)
    ? ["sup", { class: editorClassNames.noteMark }, labelOf()]
    : drawnBySchema(node);
}

/** The node specs a note's story is drawn with, its own reference mark drawn as `labelOf` answers */
export function noteNodeSpecs(labelOf: () => string): StoryNodeSpecs {
  return { rawRunContent: (node) => markSpec(node, labelOf) };
}

/**
 * The same drawing inside an editing view, with a press on the note's own number leading back to
 * the reference that calls it.
 *
 * The chip is drawn by the spec the static markup draws it by, so a note the caret goes into keeps
 * the height it was measured at. The press is answered rather than let through, so no caret ever
 * lands on the number; the chip is a preserved fragment, and the listener leaves it untouched.
 */
function noteMarkView(
  labelOf: () => string,
  back: () => void
): NodeViewConstructor {
  return (node) => {
    const { dom } = DOMSerializer.renderSpec(document, markSpec(node, labelOf));
    if (isOwnMark(node) && dom instanceof HTMLElement) {
      dom.addEventListener("mousedown", (event) => {
        event.preventDefault();
        back();
      });
    }
    return { dom };
  };
}

/**
 * Takes out of a paste what the notes part cannot be written with: an image, and a link.
 *
 * Either needs a relationship of the part it stands in, and the notes part writer writes none, so
 * both would go out naming a relationship the file does not hold. A link or an image the file
 * itself wrote inside a note stays; this is about what a paste would add.
 */
export const dropUnwritableContent: SliceNormalizer = (slice) =>
  mapSliceNodes(slice, (node) => {
    if (node.type === docxSchema.nodes.image) return null;
    const marks = node.marks.filter(
      (mark) => mark.type !== docxSchema.marks.link
    );
    return Mark.sameSet(marks, node.marks) ? node : node.mark(marks);
  });

/**
 * Takes a note reference out of a paste: a note may not hold a note.
 *
 * The lifecycle, the numbering and the writer all walk the body alone, so a reference inside a
 * story would call a note nothing numbers, counts or settles, and one copied from beside the note
 * it was pasted into would have that note calling itself. The text it stood in stays, the way the
 * other anchors leave theirs (`editor/clipboard/normalizers`).
 */
export const dropNoteReferences: SliceNormalizer = (slice, { move }) =>
  move
    ? slice
    : mapSliceNodes(slice, (node) =>
        node.type === docxSchema.nodes.noteReference ? null : node
      );

/**
 * A note takes nothing beyond character and paragraph formatting: a link and an image each need a
 * relationship of their part, which the notes part writer writes none of; a table, a list and a
 * comment need more of the package than that writer puts together; and a note may not hold a note.
 * What the file itself wrote inside a note is kept; this is about what an edit may add.
 */
const NOTE_TAKES = NO_CAPABILITY;

const NOTE_NORMALIZERS: readonly SliceNormalizer[] = [
  dropSourceIdentity,
  detachAnchors,
  dropNoteReferences,
  dropUnwritableContent,
  rekeyNumbering,
  rederiveDisplay,
];

function runOn(main: EditorView, command: Command): boolean {
  return command(main.state, (tr) => main.dispatch(tr), main);
}

/** Where the reference calling the note this key names stands, and null where the body calls none */
function referenceOf(main: EditorView, key: StoryKey) {
  const note = noteKeyOf(key);
  return note === null
    ? null
    : noteReferenceAt(main.state.doc, note.kind, note.id);
}

/**
 * Puts the caret back just after the reference that calls this note and takes the body's focus,
 * which is what leaves the note: Escape and a press on the note's own number both run it.
 */
export function returnToReference(main: EditorView, key: StoryKey): void {
  const found = referenceOf(main, key);
  if (found !== null) {
    const after = found.pos + found.node.nodeSize;
    main.dispatch(
      main.state.tr
        .setSelection(TextSelection.near(main.state.doc.resolve(after)))
        .scrollIntoView()
    );
  }
  main.focus();
}

/** How much room the number the note opens with takes, which the caret may stand after */
function leadingChipSize(story: PMNode): number {
  const first = story.firstChild?.firstChild ?? null;
  return first !== null && isOwnMark(first) ? first.nodeSize : 0;
}

/** The note's own number as the story holds it, and where it stands */
interface PlacedMark {
  readonly pos: number;
  readonly node: PMNode;
}

/** The note's own number, wherever in the story it stands, and null for a story holding none */
function ownMarkIn(story: PMNode): PlacedMark | null {
  let found: PlacedMark | null = null;
  story.descendants((node, pos) => {
    if (found === null && isOwnMark(node)) found = { pos, node };
    return found === null;
  });
  return found;
}

/**
 * The first place in the story a caret may stand, which is after the number the note is drawn by,
 * and null while the story holds no number of its own to stand after.
 */
function caretFloor(story: PMNode): number | null {
  const held = ownMarkIn(story);
  return held === null || held.pos !== markHome(story)
    ? null
    : held.pos + held.node.nodeSize;
}

/**
 * The head of the first stretch of the story a run's own content may stand in, and null for a
 * story left holding no such stretch.
 */
function markHome(story: PMNode): number | null {
  const type = docxSchema.nodes.rawRunContent;
  let at: number | null = null;
  story.descendants((node, pos) => {
    if (at !== null) return false;
    if (!node.isTextblock) return true;
    if (node.type.contentMatch.matchType(type) !== null) at = pos + 1;
    return false;
  });
  return at;
}

/**
 * Keeps the number a note is drawn by the first thing the note holds: an edit that carried it off
 * has it put back, and the caret never stands ahead of it.
 *
 * Word draws a note's number from the mark its entry opens with (`w:footnoteRef`), which arrives
 * as a preserved chip no deletion guard answers for (`docx/importPolicy`), so selecting the whole
 * of a note and deleting it took the number along with the text. It goes back at the head of the
 * first paragraph that takes it, as the very node it stood in the story as, so an emptied note
 * still goes out carrying the element Word reads. It is put back rather than the edit refused
 * because the guard the chip would need answers for a whole stretch of an edit, and a refusal
 * would leave the selected text standing too; and it is an appended transaction, as a comment a
 * body edit swept away is put back (`editor/plugins/commentRestoration`), so the host writes the
 * edit and the number as one story change and one undo takes both back. Inserting the number
 * beside the composed text rather than rewriting the node it stands in is what keeps a composition
 * over it open (`e2e/notesEditing.spec.ts`).
 *
 * Holding the caret after the number is what keeps this a rule a reader never runs into: every way
 * into a note leaves the caret where the note's own text begins, so the move above answers only
 * what a paste or a plugin wrote straight into the story. A selection reaching over the number is
 * left alone: it is an edit like any other, and what it sweeps away is put back, which keeps a
 * paste over the whole of a note one paragraph.
 */
function ownMarkFirst(): Plugin {
  return new Plugin({
    appendTransaction(transactions, oldState, newState) {
      const tr = newState.tr;
      const home = markHome(newState.doc);
      if (transactions.some((changed) => changed.docChanged) && home !== null) {
        const held = ownMarkIn(newState.doc);
        const lost = ownMarkIn(oldState.doc);
        if (held === null) {
          if (lost !== null) tr.insert(home, lost.node);
        } else if (held.pos !== home) {
          tr.delete(held.pos, held.pos + held.node.nodeSize).insert(
            home,
            held.node
          );
        }
      }
      const floor = caretFloor(tr.doc);
      if (floor !== null && tr.selection.empty && tr.selection.from < floor) {
        tr.setSelection(
          TextSelection.create(tr.doc, Math.min(floor, tr.doc.content.size))
        );
      }
      return tr.docChanged || tr.selectionSet ? tr : null;
    },
  });
}

/**
 * Whether the note holds nothing but the number it opens with, judged by content rather than by
 * text, since an image or an empty table spells no text at all.
 */
function holdsNothing(story: PMNode): boolean {
  const only = story.childCount === 1 ? story.firstChild : null;
  return (
    only !== null &&
    only.type === docxSchema.nodes.paragraph &&
    only.content.size === leadingChipSize(story)
  );
}

/**
 * Deletes an empty note and the reference that calls it, and hands the caret back to where the
 * reference stood. Only the reference is deleted: the note goes with its last reference
 * (`editor/plugins/noteLifecycle`), in the same history event, so one undo brings both back.
 */
function deleteEmptyNote(main: EditorView, key: StoryKey): Command {
  return (state, dispatch) => {
    if (!state.selection.empty) return false;
    if (!holdsNothing(state.doc)) return false;
    const $at = state.selection.$from;
    if ($at.index(0) !== 0 || $at.parentOffset > leadingChipSize(state.doc)) {
      return false;
    }
    const found = referenceOf(main, key);
    if (found === null) return false;
    const tr = main.state.tr.delete(found.pos, found.pos + found.node.nodeSize);
    if (!transactionAllowed(tr, main.state)) return false;
    if (dispatch) {
      main.dispatch(
        tr
          .setSelection(TextSelection.near(tr.doc.resolve(found.pos)))
          .scrollIntoView()
      );
      main.focus();
    }
    return true;
  };
}

/**
 * The host a note's editing view writes through, whichever kind of note stands under the key.
 *
 * Every edit leaves as that kind's own body command, so a lock around the reference and the
 * editing protection judge a note edit where they judge a body edit, and the main document keeps
 * the one history both share.
 */
export function noteHost(
  main: EditorView,
  activate: StoryHost["activate"]
): StoryHost {
  return {
    state: () => main.state,
    write(key, story) {
      const note = noteKeyOf(key);
      return (
        note !== null && runOn(main, noteBodyCommand(note.kind, note.id, story))
      );
    },
    undo: () => runOn(main, undo),
    redo: () => runOn(main, redo),
    leave: (key) => returnToReference(main, key),
    shut(key) {
      if (editsShut(main.state)) return true;
      const found = referenceOf(main, key);
      if (found === null) return true;
      return editShut(main.state, {
        kind: "replace",
        from: found.pos,
        to: found.pos + found.node.nodeSize,
      });
    },
    activate,
  };
}

/**
 * What a note adds to the story view: its chip, the number it keeps first, its key rule, and what
 * a paste may bring in
 */
export function noteExtensions(
  main: EditorView,
  key: StoryKey,
  labelOf: () => string
): StoryExtensions {
  return {
    plugins: [
      ownMarkFirst(),
      keymap({ Backspace: deleteEmptyNote(main, key) }),
    ],
    nodeViews: {
      rawRunContent: noteMarkView(labelOf, () => returnToReference(main, key)),
    },
    normalizers: NOTE_NORMALIZERS,
    takes: NOTE_TAKES,
  };
}
