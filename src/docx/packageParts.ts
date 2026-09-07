/**
 * Where a part sits in the package, asked one way by every reader and writer.
 *
 * A part beside the body is found through the relationship the main part declares for it, and a
 * part the export adds takes the first name beside the main part that nothing has claimed. Each
 * reader used to find its part with a lookup of its own, so the rule for what counts as the
 * related part lived in five places.
 */

import { decodeUtf8 } from "../ooxml/xml";
import {
  directoryOf,
  readRelationships,
  relsPathOf,
  resolveTarget,
} from "./relationships";

export const CONTENT_TYPES_PATH = "[Content_Types].xml";

/**
 * Where the part the main part relates under this type sits. null for no such relationship, and
 * for one that points outside the package, which no part of it can stand behind.
 */
export function relatedPartPath(
  parts: ReadonlyMap<string, Uint8Array>,
  mainPartPath: string,
  type: string
): string | null {
  const relationship = readRelationships(parts, relsPathOf(mainPartPath)).find(
    (entry) => entry.type === type && !entry.external
  );
  return relationship === undefined
    ? null
    : resolveTarget(mainPartPath, relationship.target);
}

/** A part as text, and null for a path there is no part at, so a missing relationship reads as a missing part */
export function readPart(
  parts: ReadonlyMap<string, Uint8Array>,
  path: string | null
): string | null {
  const bytes = path === null ? undefined : parts.get(path);
  return bytes ? decodeUtf8(bytes).text : null;
}

/** `comments.xml`, then `comments2.xml`, and so on beside the main part: the first name no part of the package has taken */
export function availablePartPath(
  parts: ReadonlyMap<string, Uint8Array>,
  mainPartPath: string,
  stem: string
): string {
  const directory = directoryOf(mainPartPath);
  for (let suffix = 0; ; suffix += 1) {
    const path = `${directory}${stem}${suffix === 0 ? "" : suffix + 1}.xml`;
    if (!parts.has(path)) return path;
  }
}
