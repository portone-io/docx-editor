/**
 * The note an edit puts into a document: a new, empty one, and a copy of one the document holds.
 *
 * Both are side stories (`docx/story`), written back by the notes part writer (`./writing`) as an
 * entry of their own, which is why a copy names none of the blocks its original was read from.
 */

import { Fragment, type Node as PMNode } from "prosemirror-model";
import { elementXml } from "../../ooxml/element";
import { wName } from "../../ooxml/names";
import { docxSchema } from "../../schema";
import {
  NOTE_KINDS,
  type NoteKind,
  type StoryKey,
  storiesOf,
  storyKey,
} from "../../schema/stories";
import { CLONE_POLICIES } from "../cloning";
import type { FormattingContext } from "../formatting";
import { originalBlock, type SessionStore } from "../session";
import { withStyleFormats } from "../story";

/** The styles Word writes a note of one kind in, and the element that stands for its number inside it */
interface NoteMarkup {
  /** The paragraph style of the note's text */
  readonly textStyle: string;
  /** The character style of the note's number, where the text calls the note and inside the note */
  readonly referenceStyle: string;
  readonly numberElement: string;
}

const NOTE_MARKUP: Readonly<Record<NoteKind, NoteMarkup>> = {
  footnote: {
    textStyle: "FootnoteText",
    referenceStyle: "FootnoteReference",
    numberElement: "footnoteRef",
  },
  endnote: {
    textStyle: "EndnoteText",
    referenceStyle: "EndnoteReference",
    numberElement: "endnoteRef",
  },
};

/**
 * The elements that stand for a note's own number inside its entry, which is what Word draws the
 * number from and what a note that keeps its number keeps (`editor/notes/noteSurface`).
 *
 * One per note kind, read off the kinds rather than restated, so a kind added to `NOTE_KINDS`
 * cannot arrive without the element its number is written as: the table above is keyed by the
 * kind, and a missing row is a compile error.
 */
export const NOTE_NUMBER_ELEMENTS: readonly string[] = NOTE_KINDS.map(
  (kind) => NOTE_MARKUP[kind].numberElement
);

/**
 * The run properties a note's number is written with: the document's own style for it where the
 * document defines one, and superscript where it does not.
 */
export function noteNumberRunProps(
  kind: NoteKind,
  formatting: FormattingContext
): string {
  const { referenceStyle } = NOTE_MARKUP[kind];
  const property =
    formatting.styles.get(referenceStyle)?.type === "character"
      ? elementXml(wName("rStyle"), [[wName("val"), referenceStyle]])
      : elementXml(wName("vertAlign"), [[wName("val"), "superscript"]]);
  return elementXml(wName("rPr"), [], [property]);
}

function noteTextProps(
  kind: NoteKind,
  formatting: FormattingContext
): string | null {
  const { textStyle } = NOTE_MARKUP[kind];
  return formatting.styles.get(textStyle)?.type === "paragraph"
    ? elementXml(
        wName("pPr"),
        [],
        [elementXml(wName("pStyle"), [[wName("val"), textStyle]])]
      )
    : null;
}

/**
 * The story of a new note holding no text yet: one paragraph in the document's style for note
 * text, opening with the note's number the way Word writes a note it inserts.
 */
export function newNoteStory(
  kind: NoteKind,
  formatting: FormattingContext
): PMNode {
  const { numberElement } = NOTE_MARKUP[kind];
  const number = docxSchema.nodes.rawRunContent.create(
    {
      xml: elementXml(wName(numberElement), []),
      element: numberElement,
      display: "chip",
      text: null,
      guarded: false,
    },
    null,
    [
      docxSchema.marks.run.create({
        rPr: noteNumberRunProps(kind, formatting),
      }),
    ]
  );
  const paragraph = docxSchema.nodes.paragraph.create(
    { pPr: noteTextProps(kind, formatting) },
    number
  );
  return withStyleFormats(
    docxSchema.nodes.doc.create(null, [paragraph]),
    formatting
  );
}

const DECIMAL = /^-?\d+$/;

/**
 * Every id a note of this kind already answers to: the entries the opened part holds, the stories
 * the document holds, and the ids the references name, whether or not a note stands behind them.
 */
export function takenNoteIds(
  doc: PMNode,
  kind: NoteKind,
  reserved: ReadonlySet<StoryKey>
): Set<string> {
  const prefix = storyKey(kind, "");
  const ids = new Set(
    [...reserved, ...Object.keys(storiesOf(doc))].flatMap((key) =>
      key.startsWith(prefix) ? [key.slice(prefix.length)] : []
    )
  );
  doc.descendants((node) => {
    const id: unknown = node.attrs.id;
    if (
      node.type === docxSchema.nodes.noteReference &&
      node.attrs.kind === kind &&
      typeof id === "string"
    ) {
      ids.add(id);
    }
    return true;
  });
  return ids;
}

/** The id one above the largest of these, and never below one */
export function nextNoteId(taken: Iterable<string>): string {
  const largest = Array.from(taken)
    .filter((id) => DECIMAL.test(id))
    .map((id) => BigInt(id))
    .reduce((max, id) => (id > max ? id : max), 0n);
  return (largest + 1n).toString();
}

function copiedBlock(node: PMNode, session: SessionStore | null): PMNode {
  if (node.isInline) return node;
  if (node.type === docxSchema.nodes.paragraph) {
    return node.type.create(
      CLONE_POLICIES.paragraph.attrs(node, "copy"),
      node.content,
      node.marks
    );
  }
  if (node.type === docxSchema.nodes.rawBlock) {
    const held: unknown = node.attrs.xml;
    const xml =
      typeof held === "string"
        ? held
        : session === null
          ? undefined
          : originalBlock(node, session)?.xml;
    return xml === undefined
      ? node
      : node.type.create({ ...node.attrs, xml, srcId: null }, null, node.marks);
  }
  const attrs =
    node.attrs.srcId === undefined || node.attrs.srcId === null
      ? node.attrs
      : { ...node.attrs, srcId: null };
  return node.type.create(
    attrs,
    Fragment.fromArray(
      node.children.map((child) => copiedBlock(child, session))
    ),
    node.marks
  );
}

/**
 * The story a copied reference's note is given: the same content, naming none of the blocks the
 * original was read from and none of its paragraph ids, so the export writes it as an entry of its
 * own rather than as the bytes of the original a second time (`docx/cloning`). A preserved block
 * that only names its fragment takes the fragment's XML along.
 */
export function copiedNoteStory(
  story: PMNode,
  session: SessionStore | null
): PMNode {
  return copiedBlock(story, session);
}
