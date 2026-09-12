/**
 * Writes an edited, added, or removed note back into the part of its kind, one entry per story
 * (`docx/storyParts`).
 *
 * The two parts differ in nothing but the names they are written under, so both are laid out by
 * one description read off the kind: a `w:footnotes` part holds a `w:footnote` apiece and a
 * `w:endnotes` part a `w:endnote` (§17.11.15, §17.11.16).
 *
 * A part this writer creates opens with the separator and continuation-separator entries Word lays
 * the line above the notes out with (§17.11.1, §17.11.23). `settings.xml` is left alone: the
 * `w:footnote` and `w:endnote` references there are optional (`CT_FtnDocProps`, `CT_EdnDocProps`),
 * and a Google Docs export writes none.
 */

import type { Node as PMNode } from "prosemirror-model";
import { elementXml } from "../../ooxml/element";
import { wName } from "../../ooxml/names";
import { R_NS } from "../../ooxml/xml";
import type { NoteKind, StoryKey } from "../../schema/stories";
import type { PartPlanner } from "../partPlan";
import type { SessionStore } from "../session";
import { type StoryEntriesPart, storyEntriesPlanner } from "../storyParts";

const NOTE_CONTENT_TYPES: Readonly<Record<NoteKind, string>> = {
  footnote:
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  endnote:
    "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
};

/** The special entries a new notes part opens with, each under the id Word gives it */
const SEPARATORS: readonly { readonly type: string; readonly id: number }[] = [
  { type: "separator", id: -1 },
  { type: "continuationSeparator", id: 0 },
];

const DECIMAL = /^-?\d+$/;

/**
 * The id each separator is written under: the one Word gives it where nothing holds that yet, and
 * otherwise the id one below the smallest held, which nothing can hold.
 */
function separatorIds(taken: ReadonlySet<string>): readonly number[] {
  const held = Array.from(taken)
    .filter((id) => DECIMAL.test(id))
    .map(Number);
  return SEPARATORS.reduce<readonly number[]>((chosen, { id }) => {
    const used = [...held, ...chosen];
    return [...chosen, used.includes(id) ? Math.min(0, ...used) - 1 : id];
  }, []);
}

/** One separator entry as Word writes it, its paragraph kept clear of the spacing the styles give text */
function separatorXml(kind: NoteKind, type: string, id: number): string {
  const paragraph = elementXml(
    wName("p"),
    [],
    [
      elementXml(
        wName("pPr"),
        [],
        [
          elementXml(wName("spacing"), [
            [wName("after"), "0"],
            [wName("line"), "240"],
            [wName("lineRule"), "auto"],
          ]),
        ]
      ),
      elementXml(wName("r"), [], [elementXml(wName(type), [])]),
    ]
  );
  return elementXml(
    wName(kind),
    [
      [wName("type"), type],
      [wName("id"), String(id)],
    ],
    [paragraph]
  );
}

function separatorsXml(kind: NoteKind, taken: ReadonlySet<string>): string {
  const ids = separatorIds(taken);
  return SEPARATORS.map(({ type }, at) =>
    separatorXml(kind, type, ids[at] ?? 0)
  ).join("");
}

/**
 * The entries of this kind the part lays the page out with rather than numbers - a separator, a
 * continuation notice - which go back out as they arrived whatever the document says of them.
 *
 * The session holds the two kinds' special entries in one set, so each part takes its own out of
 * it: a footnote and an endnote can stand under the same id.
 */
function frozenNotes(
  session: SessionStore,
  kind: NoteKind
): ReadonlySet<StoryKey> {
  return new Set(
    Array.from(session.specialNotes).filter((key) => key.startsWith(`${kind}:`))
  );
}

/** The ids the references of one kind name, an orphan's included */
function referenceIds(doc: PMNode, kind: NoteKind): ReadonlySet<string> {
  const ids = new Set<string>();
  doc.descendants((node) => {
    const id: unknown = node.attrs.id;
    if (
      node.type.name === "noteReference" &&
      node.attrs.kind === kind &&
      typeof id === "string"
    ) {
      ids.add(id);
    }
    return true;
  });
  return ids;
}

/** The part one kind of note is written into, which every name it takes is the kind's own plural */
function notesPart(kind: NoteKind): StoryEntriesPart {
  const plural = `${kind}s`;
  return {
    name: plural,
    kind,
    relType: `${R_NS}/${plural}`,
    contentType: NOTE_CONTENT_TYPES[kind],
    stem: plural,
    root: plural,
    entry: kind,
    referencedIds: (doc) => referenceIds(doc, kind),
    frozenEntries: (session) => frozenNotes(session, kind),
    prelude: (taken) => separatorsXml(kind, taken),
  };
}

export const FOOTNOTES_PART: StoryEntriesPart = notesPart("footnote");

export const ENDNOTES_PART: StoryEntriesPart = notesPart("endnote");

export const footnotesPlanner: PartPlanner =
  storyEntriesPlanner(FOOTNOTES_PART);

export const endnotesPlanner: PartPlanner = storyEntriesPlanner(ENDNOTES_PART);
