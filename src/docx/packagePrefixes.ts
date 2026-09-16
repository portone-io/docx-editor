/**
 * Spelling a whole package under the prefixes this editor writes, as it is opened.
 * `spec/notes/conformance.md` holds the decision and what it rests on.
 *
 * Every reader and every writer re-reads its part out of the one map the session holds, so the
 * rewrite belongs to that map rather than to any one of them: the main part, the comments, the
 * notes, each header and footer and the numbering then all arrive spelled alike.
 *
 * Which parts hold markup is `docx/packageParts`' answer, the one the export reads the parts it
 * wrote back by. A part that comes back unrewritten is never re-encoded.
 */

import { withEditorPrefixes } from "../ooxml/partPrefixes";
import { decodeUtf8Strictly, encodeUtf8 } from "../ooxml/xml";
import { CONTENT_TYPES_PATH, markupParts } from "./packageParts";

export function normalizePackagePrefixes(parts: Map<string, Uint8Array>): void {
  const markup = markupParts(parts.keys(), parts.get(CONTENT_TYPES_PATH));
  for (const [path, bytes] of parts) {
    if (!markup.has(path)) continue;
    const decoded = decodeUtf8Strictly(bytes);
    if (decoded === null) continue;
    const rewritten = withEditorPrefixes(decoded.text);
    if (rewritten !== decoded.text)
      parts.set(path, encodeUtf8(rewritten, decoded.hadBom));
  }
}
