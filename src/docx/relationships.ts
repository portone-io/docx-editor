/**
 * The relationships part beside the body: what it already declares, and what an export adds.
 *
 * An addition is spliced in ahead of the closing tag and the original text is left exactly as it
 * came, so a package the editor only added to keeps every byte it arrived with.
 *
 * Both writers that need a relationship come through here, and they share the ids: an image points
 * at a part inside the package (`docx/media`), a hyperlink at an address outside it, which is what
 * `TargetMode="External"` says (`docx/hyperlink`). Two writers handing out ids of their own would
 * hand out the same one.
 */

import { DocxExportError } from "../ooxml/errors";
import {
  decodeUtf8,
  elementChildren,
  encodeUtf8,
  escapeXml,
  parseXml,
} from "../ooxml/xml";

const RELATIONSHIPS_NS =
  "http://schemas.openxmlformats.org/package/2006/relationships";

export function directoryOf(partPath: string): string {
  return partPath.replace(/[^/]+$/, "");
}

export function relsPathOf(partPath: string): string {
  const directory = directoryOf(partPath);
  return `${directory}_rels/${partPath.slice(directory.length)}.rels`;
}

/** A relationship target resolved against the part that declares it */
export function resolveTarget(partPath: string, target: string): string {
  const combined = target.startsWith("/")
    ? target.slice(1)
    : directoryOf(partPath) + target;
  const segments: string[] = [];
  for (const segment of combined.split("/")) {
    if (!segment || segment === ".") continue;
    if (segment === "..") {
      segments.pop();
      continue;
    }
    segments.push(segment);
  }
  return segments.join("/");
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  /** Set when the target is a file outside the package */
  external: boolean;
}

export function readRelationships(
  parts: Map<string, Uint8Array>,
  relsPath: string
): Relationship[] {
  const bytes = parts.get(relsPath);
  if (!bytes) return [];
  return elementChildren(parseXml(decodeUtf8(bytes).text).documentElement)
    .filter((el) => el.localName === "Relationship")
    .map((el) => ({
      id: el.getAttribute("Id") ?? "",
      type: el.getAttribute("Type") ?? "",
      target: el.getAttribute("Target") ?? "",
      external: el.getAttribute("TargetMode") === "External",
    }));
}

/** A relationship an export puts in. The target is escaped here, so hand it in as it reads */
export interface NewRelationship {
  type: string;
  target: string;
  /** Whether the target lies outside the package, which is every address a hyperlink names */
  external?: boolean;
}

export interface RelationshipWriter {
  /** Every relationship the part declared when the document was opened */
  readonly opened: readonly Relationship[];
  /** Puts one in and hands back the id it went in under */
  add(relationship: NewRelationship): string;
  /**
   * The part with everything added, built from the original text. null when nothing was added,
   * which is what leaves the part out of the export altogether.
   */
  part(original: Uint8Array | undefined): Uint8Array | null;
}

interface AddedRelationship extends NewRelationship {
  id: string;
}

/** The next `rIdN` no relationship has taken yet. The chosen id is marked as taken */
function takeRelId(taken: Set<string>): string {
  for (let number = taken.size + 1; ; number += 1) {
    const id = `rId${number}`;
    if (!taken.has(id)) {
      taken.add(id);
      return id;
    }
  }
}

function relationshipXml(added: AddedRelationship): string {
  const mode = added.external ? ' TargetMode="External"' : "";
  return `<Relationship Id="${added.id}" Type="${added.type}" Target="${escapeXml(added.target)}"${mode}/>`;
}

/** A part that is not there yet is written from scratch */
function relationshipsPart(
  original: Uint8Array | undefined,
  added: readonly AddedRelationship[]
): Uint8Array {
  const entries = added.map(relationshipXml).join("");
  if (!original) {
    return encodeUtf8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        `<Relationships xmlns="${RELATIONSHIPS_NS}">${entries}</Relationships>`,
      false
    );
  }
  const { text, hadBom } = decodeUtf8(original);
  const closing = text.lastIndexOf("</Relationships>");
  if (closing === -1) {
    throw new DocxExportError(
      "malformed-xml",
      "the relationships part has no closing tag"
    );
  }
  return encodeUtf8(
    text.slice(0, closing) + entries + text.slice(closing),
    hadBom
  );
}

export function relationshipWriter(
  opened: readonly Relationship[]
): RelationshipWriter {
  const taken = new Set(opened.map((relationship) => relationship.id));
  const added: AddedRelationship[] = [];
  return {
    opened,
    add: (relationship) => {
      const id = takeRelId(taken);
      added.push({ ...relationship, id });
      return id;
    },
    part: (original) =>
      added.length === 0 ? null : relationshipsPart(original, added),
  };
}
