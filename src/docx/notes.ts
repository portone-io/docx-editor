/** Reads footnote and endnote bodies related from the main document story. */

import {
  attributeByLocalName,
  decodeUtf8,
  elementChildren,
  parseXml,
  R_NS,
  W_NS,
} from "../ooxml/xml";
import { relatedPartPath } from "./packageParts";

export type NoteKind = "footnote" | "endnote";

export interface ImportedNote {
  kind: NoteKind;
  id: string;
  label: string;
  text: string;
  type: string | null;
}

export interface ImportedNotePart {
  partPath: string | null;
  ordered: readonly ImportedNote[];
  byId: ReadonlyMap<string, ImportedNote>;
}

export interface ImportedNotes {
  footnotes: ImportedNotePart;
  endnotes: ImportedNotePart;
}

const EMPTY_PART: ImportedNotePart = {
  partPath: null,
  ordered: [],
  byId: new Map(),
};

export const NO_NOTES: ImportedNotes = {
  footnotes: EMPTY_PART,
  endnotes: EMPTY_PART,
};

function inlineText(node: Element): string {
  if (node.localName === "t") return node.textContent ?? "";
  if (node.localName === "tab") return "\t";
  if (node.localName === "br" || node.localName === "cr") return "\n";
  if (node.localName === "noBreakHyphen") return "\u2011";
  if (node.localName === "softHyphen") return "\u00ad";
  return elementChildren(node).map(inlineText).join("");
}

function noteText(el: Element): string {
  return Array.from(el.getElementsByTagNameNS(W_NS, "p"))
    .map(inlineText)
    .join("\n");
}

function notePart(
  parts: Map<string, Uint8Array>,
  mainPartPath: string,
  kind: NoteKind
): ImportedNotePart {
  const partPath = relatedPartPath(parts, mainPartPath, `${R_NS}/${kind}s`);
  if (partPath === null) return EMPTY_PART;

  const bytes = parts.get(partPath);
  if (!bytes) return { ...EMPTY_PART, partPath };

  const root = parseXml(decodeUtf8(bytes).text).documentElement;
  const elements = elementChildren(root).filter((el) => el.localName === kind);
  const ordered = elements.flatMap((el): ImportedNote[] => {
    const id = attributeByLocalName(el, "id");
    if (id === null) return [];
    const type = attributeByLocalName(el, "type");
    const regular = type === null || type === "normal";
    const label = regular ? "?" : "";
    return [{ kind, id, label, text: noteText(el), type }];
  });
  const byId = new Map<string, ImportedNote>();
  for (const note of ordered) {
    if (!byId.has(note.id)) byId.set(note.id, note);
  }
  return { partPath, ordered, byId };
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
