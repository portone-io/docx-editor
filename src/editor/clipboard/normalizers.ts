/**
 * What a slice has to obey before it may be put into this document.
 *
 * A slice reaching a paste says where its blocks came from, which list its paragraphs belong to,
 * which comment its markers open, and what everything is drawn with. None of those mean anything
 * outside the document that wrote them: a block naming a fragment of the file would be written
 * from that fragment a second time, a list number naming no definition here would be exported
 * bare, and a comment marker copied along would open a range the comment part never opens. So a
 * paste is a normalization rather than an insertion, and each rule stands as one normalizer over
 * the whole slice, in the order they are declared.
 *
 * A move drop is the exception the first two rules read: content dragged from one place in this
 * very document to another is the same content, so it keeps the identity it had and the anchors it
 * stood in. The source of a move is deleted along with the drop, so nothing is duplicated.
 */

import {
  type Attrs,
  Fragment,
  Mark,
  type Node as PMNode,
  Slice,
} from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import {
  paragraphAttrsOf,
  resolveParagraph,
  resolveRun,
} from "../../docx/formatting";
import {
  copiedNoteStory,
  nextNoteId,
  takenNoteIds,
} from "../../docx/notes/newNote";
import { withListNumbering } from "../../docx/paraProps";
import type { NewLists } from "../../numbering/listRegistry";
import {
  allocateList,
  type ListKind,
  templateList,
} from "../../numbering/listTemplate";
import type { NewList } from "../../numbering/parseNumbering";
import {
  COMMENT_RANGE_MARKERS,
  PERMISSION_MARKERS,
  RANGE_MARKERS,
} from "../../ooxml/rangeMarkers";
import { docxSchema } from "../../schema";
import {
  EDITABLE_NOTE_KINDS,
  type NoteKey,
  type NoteKind,
  storyKey,
} from "../../schema/stories";
import { listRefOf, numIdsIn } from "../commands/listCommands";
import { documentFormatting } from "../documentStyles";
import { documentOf } from "../editorDocument";
import {
  canStartNewList,
  documentNumbering,
} from "../plugins/numberingDecorations";
import type { PastedContent } from "./htmlReader";
import {
  type ListKinds,
  NO_LIST_KINDS,
  NO_NOTE_STORIES,
  type NoteStories,
} from "./internalChannel";

/** What one normalizer answers about the document the slice is going into */
export interface NormalizeContext {
  readonly state: EditorState;
  /** A drop that moves what it carries rather than copying it, which keeps its source identity */
  readonly move: boolean;
  /**
   * The list numbers this document answers for: the ones it defines, the ones a list started
   * while editing was registered under, the ones its paragraphs wear, and the ones this very
   * paste has started.
   */
  readonly knownLists: ReadonlySet<number>;
  /**
   * What the numbers the slice names meant where it was copied, for the numbers this document
   * cannot answer for. Empty for a slice that came in as markup, whose lists were given numbers
   * of this document as they were read.
   */
  readonly listKinds: ListKinds;
  /**
   * A number for a list this slice starts, registered with the definition it is started with, as
   * the list button starts one. Null in a document with nowhere to write a definition.
   */
  startList(kind: ListKind): number | null;
  /**
   * What the notes the slice's references call said where it was copied
   * (`./internalChannel`). Empty for a slice that arrived any other way, whose references name
   * notes of a document this one knows nothing about.
   */
  readonly noteStories: NoteStories;
  /**
   * A note this paste puts in, under an id no note of this document answers to, registered in
   * `PastedContent.newStories`. It answers with the id the reference is pasted under.
   */
  startNote(kind: NoteKind, story: PMNode): string;
}

export type SliceNormalizer = (
  slice: Slice,
  context: NormalizeContext
) => Slice;

function xmlAttr(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

function withAttrs(node: PMNode, attrs: Attrs): PMNode {
  return node.type.create(attrs, node.content, node.marks);
}

/**
 * The children of one fragment rewritten, with the ones the mapping takes out left out.
 *
 * `openStart` and `openEnd` here are the depths this fragment is still open to. A node the mapping
 * takes out from under one of them carries the open path, so an empty paragraph stands in its
 * place: how a paste joins the text around it is not this mapping's to decide.
 */
function mapFragment(
  fragment: Fragment,
  openStart: number,
  openEnd: number,
  f: (node: PMNode) => PMNode | null
): Fragment {
  const mapped: PMNode[] = [];
  fragment.forEach((child, _offset, index) => {
    const first = index === 0;
    const last = index === fragment.childCount - 1;
    const inside = child.isText
      ? child
      : child.type.create(
          child.attrs,
          mapFragment(
            child.content,
            first ? openStart - 1 : 0,
            last ? openEnd - 1 : 0,
            f
          ),
          child.marks
        );
    const next = f(inside);
    if (next !== null) mapped.push(next);
    else if ((first && openStart > 0) || (last && openEnd > 0)) {
      mapped.push(docxSchema.nodes.paragraph.create());
    }
  });
  return Fragment.fromArray(mapped);
}

/** The slice with every node rewritten depth first, and the ones mapped to null taken out */
export function mapSliceNodes(
  slice: Slice,
  f: (node: PMNode) => PMNode | null
): Slice {
  return new Slice(
    mapFragment(slice.content, slice.openStart, slice.openEnd, f),
    slice.openStart,
    slice.openEnd
  );
}

/**
 * Takes the source identity off what is pasted.
 *
 * A block names the fragment of the opened file it came from, and the export writes that fragment
 * back verbatim rather than building one. Two blocks naming the same fragment would therefore go
 * out as the same bytes twice, identifiers and all, so a pasted copy names none and is written
 * from the model instead. A preserved block holding its own XML is that XML wherever it stands and
 * travels whole; one that only names a fragment has nothing of its own to travel with, so it goes.
 */
export const dropSourceIdentity: SliceNormalizer = (slice, { move }) =>
  move
    ? slice
    : mapSliceNodes(slice, (node) => {
        if (node.type === docxSchema.nodes.rawBlock) {
          return node.attrs.xml === null
            ? null
            : withAttrs(node, { ...node.attrs, srcId: null });
        }
        if (
          node.type !== docxSchema.nodes.paragraph &&
          node.type !== docxSchema.nodes.table
        ) {
          return node;
        }
        return node.attrs.srcId === null
          ? node
          : withAttrs(node, { ...node.attrs, srcId: null });
      });

/**
 * Gives a pasted list a number this document defines.
 *
 * A number the destination knows nothing about is exported as a list nothing defines, which the
 * export invariants refuse. The paragraphs that shared such a number still share the one they are
 * given, so a pasted list stays one list, started the way the list button starts one and drawn and
 * written from the destination's own definition from here on.
 *
 * Which of the two kinds it is started as is the one thing the number alone does not say. A copy
 * made in this same session was remembered alongside what its numbers meant there
 * (`./internalChannel`), so a bulleted list comes back bulleted; a number arriving any other way
 * has no kind to be read off it and counts.
 */
export const rekeyNumbering: SliceNormalizer = (slice, context) => {
  const given = new Map<number, number | null>();
  return mapSliceNodes(slice, (node) => {
    if (node.type !== docxSchema.nodes.paragraph) return node;
    const ref = listRefOf(node);
    if (ref === null || context.knownLists.has(ref.numId)) return node;
    const taken = given.has(ref.numId)
      ? (given.get(ref.numId) ?? null)
      : context.startList(context.listKinds.get(ref.numId) ?? "numbered");
    given.set(ref.numId, taken);
    const props =
      taken === null
        ? withListNumbering(node.attrs.pPr, {
            numbering: null,
            indent: { kind: "clearHanging" },
          })
        : withListNumbering(node.attrs.pPr, {
            numbering: { numId: taken, ilvl: ref.ilvl },
            indent: { kind: "keep" },
          });
    return props === null
      ? node
      : withAttrs(node, { ...node.attrs, pPr: props.pPr });
  });
};

/**
 * What each kind of node anchors outside the paragraph it stands in, and whether this one still
 * anchors it once it is pasted.
 *
 * A reference to a note of a kind no writer puts back together is one of them: nothing here can
 * write that note a second time. A footnote reference is not, since its note travels with the copy
 * and is put in beside it (`duplicateNotes`).
 */
const ANCHORS: Readonly<Record<string, (node: PMNode) => boolean>> = {
  commentStart: () => true,
  commentEnd: () => true,
  commentReference: () => true,
  noteReference: (node) =>
    !EDITABLE_NOTE_KINDS.some((kind) => kind === node.attrs.kind),
};

/**
 * The elements that call something the package writes elsewhere rather than opening a range, and
 * the number a note draws inside its own body, which stands for nothing in the text it is
 * pasted into.
 */
const REFERENCE_ELEMENTS = [
  "commentReference",
  "footnoteReference",
  "endnoteReference",
  "footnoteRef",
  "endnoteRef",
];

/**
 * The mark a note's own body opens with, which stands for the number that note is called by.
 *
 * Outside its note it names nothing: the body would draw a number the text never asked for, and
 * the writer would put a note's own mark into the document story.
 */
const NOTE_OWN_MARKS = ["footnoteRef", "endnoteRef"];

/**
 * The elements a preserved fragment holds that anchor one, kept as the XML they arrived as.
 *
 * The ranges are the ones the import keeps as hidden markers (`docx/importPolicy`), read from the
 * vocabulary they are declared in rather than restated here, so a marker the reader learns to keep
 * is a marker a paste knows to detach.
 */
const ANCHOR_ELEMENTS: ReadonlySet<string> = new Set([
  ...RANGE_MARKERS,
  ...PERMISSION_MARKERS,
  ...COMMENT_RANGE_MARKERS,
  ...REFERENCE_ELEMENTS,
  ...NOTE_OWN_MARKS,
]);

const PRESERVED_INLINE: ReadonlySet<string> = new Set([
  "rawInline",
  "rawRunContent",
]);

/** The local name of the element a preserved fragment was opened from */
function preservedName(node: PMNode): string | null {
  const xml: unknown = node.attrs.xml;
  const opening =
    typeof xml === "string" ? /^<\s*(?:[^\s:>/]+:)?([^\s>/]+)/.exec(xml) : null;
  if (opening?.[1] !== undefined) return opening[1];
  const element: unknown = node.attrs.element;
  return typeof element === "string" ? element : null;
}

/**
 * Takes the anchors off what is pasted.
 *
 * A comment marker, a bookmark and an endnote reference all point at something written elsewhere
 * in the package: the comment, the bookmark's other end, the note's body. A copy of the marker
 * alone points at the same thing a second time, which is a duplicate identifier the file may not
 * hold and a range the reader cannot close. Nothing here duplicates what they point at, so the
 * anchor goes and the text it stood in stays, which is what Word does with a note it has no body
 * for.
 *
 * A move drop duplicates nothing: the source goes as the drop lands, so the one anchor there was
 * travels with the text it opened. Taking it off would delete a marker the document has no way to
 * write back, which is a change the preserved guard refuses whole, and the drag would do nothing.
 */
export const detachAnchors: SliceNormalizer = (slice, { move }) =>
  move
    ? slice
    : mapSliceNodes(slice, (node) => {
        if (ANCHORS[node.type.name]?.(node) === true) return null;
        if (!PRESERVED_INLINE.has(node.type.name)) return node;
        const name = preservedName(node);
        return name !== null && ANCHOR_ELEMENTS.has(name) ? null : node;
      });

/**
 * Gives a pasted note reference a note of its own.
 *
 * A note is written once and called from the text by its id, so a second reference naming that id
 * is one note drawn with one number where the reader made two. The note each reference called was
 * remembered as the slice was copied (`./internalChannel`) - a cut deletes a footnote together
 * with its last reference, so the document no longer holds it by the time the paste lands - and
 * each reference is pasted under a fresh id with a copy of that note beside it.
 *
 * A reference no remembered note answers for points at nothing here: a copy from another editor or
 * another document numbers its notes against its own file, and markup naming one says nothing
 * about what it said. It goes the way the other anchors do.
 *
 * Which kinds of note this answers for is not its own to choose: `EDITABLE_NOTE_KINDS`
 * (`schema/stories`) is the one value that says so, and the copy and the paste read it three
 * places over - `ANCHORS` above has already dropped a reference of any other kind, and a copy
 * remembers the stories of the editable kinds alone (`./internalChannel`, `./htmlReader`).
 */
export const duplicateNotes: SliceNormalizer = (
  slice,
  { move, noteStories, startNote }
) =>
  move
    ? slice
    : mapSliceNodes(slice, (node) => {
        if (node.type !== docxSchema.nodes.noteReference) return node;
        const kind = EDITABLE_NOTE_KINDS.find(
          (candidate) => candidate === node.attrs.kind
        );
        const id: unknown = node.attrs.id;
        if (kind === undefined || typeof id !== "string") return null;
        const story = noteStories.get(storyKey(kind, id));
        if (story === undefined) return null;
        // The XML the reference arrived as names the note it was copied from, so the pasted one
        // is written from its attrs instead
        return withAttrs(node, {
          ...node.attrs,
          id: startNote(kind, story),
          referenceXml: null,
        });
      });

/**
 * Works out again what the pasted paragraphs are drawn with.
 *
 * The values a paragraph and the runs in it carry were resolved against the styles, the defaults
 * and the list definitions of wherever the slice came from, this document's own list definitions
 * after the numbers have just been reissued among them. Resolving them here is what makes the
 * numbering the paste registers the numbering the document is then found to wear.
 */
export const rederiveDisplay: SliceNormalizer = (slice, { state }) => {
  const formatting = documentFormatting(state);
  return mapSliceNodes(slice, (node) => {
    if (node.type !== docxSchema.nodes.paragraph) return node;
    const paragraph = resolveParagraph(xmlAttr(node.attrs.pPr), formatting);
    const inline = node.children.map((child) => {
      const marks = child.marks.map((mark) =>
        mark.type === docxSchema.marks.run
          ? mark.type.create({
              ...mark.attrs,
              format: resolveRun(
                xmlAttr(mark.attrs.rPr),
                paragraph,
                formatting
              ),
            })
          : mark
      );
      return Mark.sameSet(marks, child.marks) ? child : child.mark(marks);
    });
    return node.type.create(
      { ...node.attrs, ...paragraphAttrsOf(paragraph) },
      Fragment.fromArray(inline),
      node.marks
    );
  });
};

export const DEFAULT_NORMALIZERS: readonly SliceNormalizer[] = [
  dropSourceIdentity,
  detachAnchors,
  duplicateNotes,
  rekeyNumbering,
  rederiveDisplay,
];

/**
 * The pasted content as this document may hold it, the definitions of the lists the paste starts -
 * the ones a reading started, and the ones a reissued number was taken for - and the notes it puts
 * in beside the references calling them.
 */
export function normalizePasted(
  content: PastedContent,
  state: EditorState,
  move: boolean,
  normalizers: readonly SliceNormalizer[] = DEFAULT_NORMALIZERS
): PastedContent {
  let numbering = documentNumbering(state);
  const known = new Set([
    ...numbering.lists.keys(),
    ...numbering.added.keys(),
    ...numIdsIn(state.doc),
    ...content.newLists.keys(),
  ]);
  const started = new Map<number, NewList>();
  const document = documentOf(state);
  const takenNotes = new Map<NoteKind, Set<string>>();
  const startedNotes = new Map<NoteKey, PMNode>();
  const context: NormalizeContext = {
    state,
    move,
    knownLists: known,
    listKinds: content.listKinds ?? NO_LIST_KINDS,
    noteStories: content.noteStories ?? NO_NOTE_STORIES,
    startNote(kind, story) {
      const ids =
        takenNotes.get(kind) ??
        takenNoteIds(state.doc, kind, document.reservedNoteKeys);
      takenNotes.set(kind, ids);
      const id = nextNoteId(ids);
      ids.add(id);
      startedNotes.set(
        storyKey(kind, id),
        copiedNoteStory(story, document.session)
      );
      return id;
    },
    startList(kind) {
      if (!canStartNewList(state)) return null;
      const list = templateList(kind);
      const taken = allocateList(numbering, known, list);
      numbering = taken.numbering;
      known.add(taken.numId);
      started.set(taken.numId, list);
      return taken.numId;
    },
  };
  const slice = normalizers.reduce(
    (current, normalize) => normalize(current, context),
    content.slice
  );
  return {
    slice,
    newLists:
      started.size === 0
        ? content.newLists
        : (new Map([...content.newLists, ...started]) satisfies NewLists),
    newStories: startedNotes.size === 0 ? content.newStories : startedNotes,
  };
}
