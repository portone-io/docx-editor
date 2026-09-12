/**
 * What a footnote or an endnote draws and does differently from any other story.
 *
 * A note's story opens with a mark of its own (`w:footnoteRef`, `w:endnoteRef`), which the file
 * keeps as a preserved chip. Word draws that mark as the number the reference carries, so it is
 * drawn here as that label in superscript, the way the reference in the body is.
 *
 * The rest is what the story view asks a kind of story for (`editor/stories`): the host that
 * writes an edit into the main document through the note's own body command, and the extensions
 * that give a note its key rule, its chip, the number it keeps through an edit, and what a paste
 * into one may carry.
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
  const first = story.firstChild?.firstChild ?? null;
  return first !== null && isOwnMark(first) ? first.nodeSize : 0;
}

/** The note's own number, wherever in the story it stands, and null for a story holding none */
function ownMarkIn(story: PMNode): PMNode | null {
  let found: PMNode | null = null;
  story.descendants((node) => {
    if (isOwnMark(node)) found = node;
    return found === null;
  });
  return found;
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
 * Puts the number a note opens with back when an edit inside the note carries it off.
 *
 * Word draws a note's number from the mark its entry opens with (`w:footnoteRef`), which arrives
 * as a preserved chip no deletion guard answers for (`docx/importPolicy`), so selecting the whole
 * of a note and deleting it took the number along with the text and the entry went back out
 * without one. Emptying a note stays an ordinary edit - the text goes - while the number, which
 * is the entry's own rather than anything a reader wrote, goes back at the head of the first
 * paragraph that takes it, as the very node it stood in the story as - its preserved XML, its run
 * style and whatever wrapper it opened inside. So an entry nobody touched still goes out as its
 * own bytes, and an emptied one goes out carrying the element Word reads.
 *
 * Putting it back rather than refusing the edit is what keeps the key honest: the guard the chip
 * would need answers for a whole stretch of an edit, so a refusal would leave the selected text
 * standing too.
 *
 * It is an appended transaction, the way a comment a body edit swept away is put back
 * (`editor/plugins/commentRestoration`), so the story the host writes is the edit and the number
 * together: one story change in the document, and one undo for both.
 *
 * Backspace at the start of a note holding nothing but its number is a rule of its own and stays
 * one (`deleteEmptyNote`): it answers the key before any edit is made, so it takes the note and
 * the reference calling it away rather than leaving a story for this to answer for.
 *
 * A composition asks for no deferral of its own. The number goes in beside the composed text
 * rather than rewriting the node it stands in, which is the difference that lets a comment be put
 * back under an open composition as well (`editor/plugins/commentRestoration`), so a composition
 * that writes over the number keeps it and stays open. `e2e/notesEditing.spec.ts` holds that
 * against a real browser, which is the only place a composition can be measured.
 */
function ownMarkRestoration(): Plugin {
  return new Plugin({
    appendTransaction(transactions, oldState, newState) {
      if (!transactions.some((tr) => tr.docChanged)) return null;
      const lost = ownMarkIn(oldState.doc);
      if (lost === null || ownMarkIn(newState.doc) !== null) return null;
      const at = markHome(newState.doc);
      return at === null ? null : newState.tr.insert(at, lost);
    },
  });
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

/**
 * What a note adds to the story view: its chip, the number it keeps, its key rule, and what a
 * paste may bring in
 */
export function noteExtensions(
  main: EditorView,
  kind: NoteKind,
  key: StoryKey,
  labelOf: () => string
): StoryExtensions {
  return {
    plugins: [
      ownMarkRestoration(),
      keymap({ Backspace: deleteEmptyNote(main, kind, key) }),
    ],
    nodeViews: { rawRunContent: noteMarkView(labelOf) },
    normalizers: NOTE_NORMALIZERS,
    takes: NOTE_TAKES,
  };
}
