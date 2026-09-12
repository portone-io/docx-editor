/** Reads footnote and endnote bodies related from the main document story. */

import { isNumberFormat, type NumberFormat } from "../../numbering/spellers";
import { ST_DecimalNumber } from "../../ooxml/simpleTypes";
import {
  attributeByLocalName,
  childByLocalName,
  decodeUtf8,
  elementChildren,
  parseXml,
  R_NS,
} from "../../ooxml/xml";
import { type NoteKind, type StoryKey, storyKey } from "../../schema/stories";
import { relatedPartPath } from "../packageParts";

/**
 * One entry of a notes part as it arrived. What the note says is a story of its own, held on the
 * document node under `footnote:<id>` or `endnote:<id>` (`docx/story`).
 */
export interface ImportedNote {
  kind: NoteKind;
  id: string;
  type: string | null;
}

export interface ImportedNotePart {
  partPath: string | null;
  /** The raw text of the part, which is what the note stories are sliced out of. null for none */
  xml: string | null;
  ordered: readonly ImportedNote[];
  byId: ReadonlyMap<string, ImportedNote>;
}

export interface ImportedNotes {
  footnotes: ImportedNotePart;
  endnotes: ImportedNotePart;
}

const EMPTY_PART: ImportedNotePart = {
  partPath: null,
  xml: null,
  ordered: [],
  byId: new Map(),
};

function notePart(
  parts: Map<string, Uint8Array>,
  mainPartPath: string,
  kind: NoteKind
): ImportedNotePart {
  const partPath = relatedPartPath(parts, mainPartPath, `${R_NS}/${kind}s`);
  if (partPath === null) return EMPTY_PART;

  const bytes = parts.get(partPath);
  if (!bytes) return { ...EMPTY_PART, partPath };

  const xml = decodeUtf8(bytes).text;
  const root = parseXml(xml).documentElement;
  const elements = elementChildren(root).filter((el) => el.localName === kind);
  const ordered = elements.flatMap((el): ImportedNote[] => {
    const id = attributeByLocalName(el, "id");
    if (id === null) return [];
    return [{ kind, id, type: attributeByLocalName(el, "type") }];
  });
  const byId = new Map<string, ImportedNote>();
  for (const note of ordered) {
    if (!byId.has(note.id)) byId.set(note.id, note);
  }
  return { partPath, xml, ordered, byId };
}

/** Reads both note parts without modifying their original package bytes. */
export function readNotes(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): ImportedNotes {
  return {
    footnotes: notePart(parts, mainPartPath, "footnote"),
    endnotes: notePart(parts, mainPartPath, "endnote"),
  };
}

function isSpecial(note: ImportedNote): boolean {
  return note.type !== null && note.type !== "normal";
}

/**
 * The entries a notes part lays the page out with rather than numbers - a separator, a
 * continuation separator, a continuation notice (§17.11.1, §17.11.23) - as the keys their stories
 * stand under.
 */
export function specialNotesOf(notes: ImportedNotes): ReadonlySet<StoryKey> {
  return new Set(
    [...notes.footnotes.byId.values(), ...notes.endnotes.byId.values()]
      .filter(isSpecial)
      .map((note) => storyKey(note.kind, note.id))
  );
}

/** Where the count of one kind of note goes back to its start (§17.18.74) */
export type NoteRestart = "continuous" | "eachSect" | "eachPage";

/** How one kind of note is counted (§17.11.17-20) */
export interface NoteNumberingProps {
  readonly format: NumberFormat;
  readonly start: number;
  readonly restart: NoteRestart;
}

export type NoteNumbering = Readonly<Record<NoteKind, NoteNumberingProps>>;

/**
 * An omitted format is decimal for both kinds (§17.11.17, §17.11.18), an omitted start is one
 * (§17.11.20), and an omitted restart is continuous (§17.11.19).
 */
const UNNAMED_PROPS: NoteNumberingProps = {
  format: "decimal",
  start: 1,
  restart: "continuous",
};

export const DEFAULT_NOTE_NUMBERING: NoteNumbering = {
  footnote: UNNAMED_PROPS,
  endnote: UNNAMED_PROPS,
};

const RESTARTS: readonly NoteRestart[] = ["continuous", "eachSect", "eachPage"];

function valOf(props: Element, name: string): string | null {
  const child = childByLocalName(props, name);
  return child === null ? null : attributeByLocalName(child, "val");
}

/**
 * What one `w:footnotePr` or `w:endnotePr` names, leaving out every property it does not, and null
 * for no element at all.
 *
 * A format this editor has no speller for is counted in decimal, as a list marker is.
 */
export function readNoteProps(
  props: Element | null
): Partial<NoteNumberingProps> | null {
  if (props === null) return null;
  const format = valOf(props, "numFmt");
  const start = ST_DecimalNumber.parse(valOf(props, "numStart"));
  const restart = RESTARTS.find(
    (known) => known === valOf(props, "numRestart")
  );
  return {
    ...(format === null
      ? {}
      : { format: isNumberFormat(format) ? format : "decimal" }),
    ...(start === null ? {} : { start }),
    ...(restart === undefined ? {} : { restart }),
  };
}

/** The properties with whatever an override names laid over them (§17.11.5, §17.11.11) */
export function withNoteProps(
  base: NoteNumberingProps,
  override: Partial<NoteNumberingProps> | null
): NoteNumberingProps {
  return override === null ? base : { ...base, ...override };
}

/** The document-wide note properties settings.xml names (§17.11.4, §17.11.12) */
export function readNoteNumbering(settings: Document | null): NoteNumbering {
  const root = settings?.documentElement ?? null;
  const named = (name: string) =>
    readNoteProps(root === null ? null : childByLocalName(root, name));
  return {
    footnote: withNoteProps(UNNAMED_PROPS, named("footnotePr")),
    endnote: withNoteProps(UNNAMED_PROPS, named("endnotePr")),
  };
}
