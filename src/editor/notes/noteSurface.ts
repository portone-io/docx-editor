/**
 * What a footnote or an endnote draws and does differently from any other story.
 *
 * A note's story opens with a mark of its own (`w:footnoteRef`, `w:endnoteRef`), which the file
 * keeps as a preserved chip. Word draws that mark as the number the reference carries, so it is
 * drawn here as that label in superscript, the way the reference in the body is.
 *
 * The rest is what the story view asks a kind of story for (`editor/stories`): the host that
 * writes an edit into the main document through the note's own body command, and the extensions
 * that give a note its key rule, its chip, and what a paste into one may carry.
 */

import { keymap } from "prosemirror-keymap";
import {
  type DOMOutputSpec,
  DOMSerializer,
  Mark,
  type Node as PMNode,
} from "prosemirror-model";
import { type Command, TextSelection } from "prosemirror-state";
import type { EditorView, NodeViewConstructor } from "prosemirror-view";
import { docxSchema } from "../../schema";
import { editShut, transactionAllowed } from "../../schema/guards";
import { editsShut } from "../../schema/protectionState";
import { type NoteKind, type StoryKey, storyKey } from "../../schema/stories";
import { editorClassNames } from "../../styles/classNames";
import {
  detachAnchors,
  dropSourceIdentity,
  mapSliceNodes,
  rederiveDisplay,
  rekeyNumbering,
  type SliceNormalizer,
} from "../clipboard/normalizers";
import { noteBodyCommand } from "../commands/footnoteCommands";
import { redo, undo } from "../commands/historyCommands";
import { noteReferenceAt } from "../plugins/noteNavigation";
import type { StoryNodeSpecs } from "../stories/storyMarkup";
import {
  NO_CAPABILITY,
  type StoryExtensions,
  type StoryHost,
} from "../stories/storyView";

const OWN_REFERENCE_MARKS: readonly unknown[] = ["footnoteRef", "endnoteRef"];

function drawnBySchema(node: PMNode): DOMOutputSpec {
  return docxSchema.nodes.rawRunContent.spec.toDOM?.(node) ?? ["span"];
}

function markSpec(node: PMNode, labelOf: () => string): DOMOutputSpec {
  return OWN_REFERENCE_MARKS.includes(node.attrs.element)
    ? ["sup", { class: editorClassNames.noteMark }, labelOf()]
    : drawnBySchema(node);
}

/** The node specs a note's story is drawn with, its own reference mark drawn as `labelOf` answers */
export function noteNodeSpecs(labelOf: () => string): StoryNodeSpecs {
  return { rawRunContent: (node) => markSpec(node, labelOf) };
}

/**
 * The same drawing inside an editing view.
 *
 * A note the caret goes into has to keep the height it was measured at, so the chip is drawn by
 * the very spec the static markup draws it by rather than by the schema's own, which says nothing
 * about a label.
 */
function noteMarkView(labelOf: () => string): NodeViewConstructor {
  return (node) => {
    const { dom } = DOMSerializer.renderSpec(document, markSpec(node, labelOf));
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
 * Takes a note reference out of a paste.
 *
 * A note may not hold a note (5.3 of the notes plan). A reference written into a story would call
 * a note the body never calls, which nothing here numbers, counts or settles: the lifecycle, the
 * numbering and the writer all walk the body alone, and a reference copied from beside the note it
 * was pasted into would have that note calling itself. The text it stood in stays, the way the
 * other anchors leave theirs (`editor/clipboard/normalizers`).
 */
export const dropNoteReferences: SliceNormalizer = (slice, { move }) =>
  move
    ? slice
    : mapSliceNodes(slice, (node) =>
        node.type === docxSchema.nodes.noteReference ? null : node
      );

/**
 * What a note takes beyond character and paragraph formatting, which is nothing (5.3 of the notes
 * plan).
 *
 * A link and an image each name a relationship of the part they stand in and the notes part writer
 * writes none; a table, a list and a comment need more of the package than that writer puts
 * together; and a note may not hold a note. What the file itself wrote inside a note is kept and
 * its text stays editable - this is what an edit may add.
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

/** The id this key names for a note of this kind, and null for a key naming another kind */
export function noteIdIn(kind: NoteKind, key: StoryKey): string | null {
  const prefix = storyKey(kind, "");
  return key.startsWith(prefix) ? key.slice(prefix.length) : null;
}

function runOn(main: EditorView, command: Command): boolean {
  return command(main.state, (tr) => main.dispatch(tr), main);
}

/** How much room the number the note opens with takes, which the caret may stand after */
function leadingChipSize(story: PMNode): number {
  const first = story.firstChild?.firstChild;
  return first?.type === docxSchema.nodes.rawRunContent ? first.nodeSize : 0;
}

/**
 * Whether the note holds nothing but the number it opens with.
 *
 * What the note holds is judged rather than what it spells, because an image and an empty table
 * spell no text at all: nothing a reader put into a note may go under a key meant to take an
 * empty one back.
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
 * reference stood.
 *
 * Only the reference is deleted: the note goes with its last reference of its own accord
 * (`editor/plugins/noteLifecycle`), in the same history event, so one undo brings both back. A
 * note holding text is left alone, which is what keeps written words from disappearing under a
 * key meant to take a mistake back.
 */
function deleteEmptyNote(
  main: EditorView,
  kind: NoteKind,
  key: StoryKey
): Command {
  return (state, dispatch) => {
    const id = noteIdIn(kind, key);
    if (id === null || !state.selection.empty) return false;
    if (!holdsNothing(state.doc)) return false;
    const $at = state.selection.$from;
    if ($at.index(0) !== 0 || $at.parentOffset > leadingChipSize(state.doc)) {
      return false;
    }
    const found = noteReferenceAt(main.state.doc, kind, id);
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
 * The host a note's editing view writes through.
 *
 * Every edit leaves as the kind's own body command, so a lock around the reference and the
 * standing the editor runs under judge a note edit where they judge a body edit, and the main
 * document keeps the one history both share.
 */
export function noteHost(
  main: EditorView,
  kind: NoteKind,
  activate: StoryHost["activate"]
): StoryHost {
  const referenceOf = (key: StoryKey) => {
    const id = noteIdIn(kind, key);
    return id === null ? null : noteReferenceAt(main.state.doc, kind, id);
  };

  return {
    state: () => main.state,
    write(key, story) {
      const id = noteIdIn(kind, key);
      return id !== null && runOn(main, noteBodyCommand(kind, id, story));
    },
    undo: () => runOn(main, undo),
    redo: () => runOn(main, redo),
    leave(key) {
      const found = referenceOf(key);
      if (found !== null) {
        const after = found.pos + found.node.nodeSize;
        main.dispatch(
          main.state.tr
            .setSelection(TextSelection.near(main.state.doc.resolve(after)))
            .scrollIntoView()
        );
      }
      main.focus();
    },
    shut(key) {
      if (editsShut(main.state)) return true;
      const found = referenceOf(key);
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

/** What a note adds to the story view: its chip, its key rule, and what a paste may bring in */
export function noteExtensions(
  main: EditorView,
  kind: NoteKind,
  key: StoryKey,
  labelOf: () => string
): StoryExtensions {
  return {
    plugins: [keymap({ Backspace: deleteEmptyNote(main, kind, key) })],
    nodeViews: { rawRunContent: noteMarkView(labelOf) },
    normalizers: NOTE_NORMALIZERS,
    takes: NOTE_TAKES,
  };
}
