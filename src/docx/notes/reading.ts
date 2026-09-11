/** Reads footnote and endnote bodies related from the main document story. */

import {
  attributeByLocalName,
  decodeUtf8,
  elementChildren,
  parseXml,
  R_NS,
} from "../../ooxml/xml";
import { relatedPartPath } from "../packageParts";

export type NoteKind = "footnote" | "endnote";

/**
 * One entry of a notes part as it arrived. What the note says is a story of its own, held on the
 * document node under `footnote:<id>` or `endnote:<id>` (`docx/story`).
 */
export interface ImportedNote {
  kind: NoteKind;
  id: string;
  label: string;
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

export const NO_NOTES: ImportedNotes = {
  footnotes: EMPTY_PART,
  endnotes: EMPTY_PART,
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
    const type = attributeByLocalName(el, "type");
    const regular = type === null || type === "normal";
    const label = regular ? "?" : "";
    return [{ kind, id, label, type }];
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

export function noteById(
  notes: ImportedNotes,
  kind: NoteKind,
  id: string
): ImportedNote | undefined {
  return notes[`${kind}s`].byId.get(id);
}
