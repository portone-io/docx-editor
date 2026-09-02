/** Declares a part an export adds in `[Content_Types].xml`, leaving the original text as it came. */

import { DocxExportError } from "../../ooxml/errors";
import { decodeUtf8, encodeUtf8 } from "../../ooxml/xml";
import { CONTENT_TYPES_PATH } from "./constants";

const TYPES_OPEN_TAG = /<(?:[\w.-]+:)?Types\b[^>]*>/;

/**
 * The content types part with an override for this part, or null when it already declares one.
 * `current` is the part as an earlier addition in the same export left it, so that two additions
 * do not each write over the other's declaration.
 */
export function withContentType(
  parts: Map<string, Uint8Array>,
  partPath: string,
  contentType: string,
  current: Uint8Array | undefined
): Uint8Array | null {
  const original = current ?? parts.get(CONTENT_TYPES_PATH);
  if (!original) {
    throw new DocxExportError(
      "missing-content-types",
      `cannot add a part to a package that has no ${CONTENT_TYPES_PATH}`
    );
  }
  const { text, hadBom } = decodeUtf8(original);
  const partName = `/${partPath}`;
  if (
    new RegExp(
      `<(?:[\\w.-]+:)?Override[^>]+PartName=["']${partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}["']`,
      "i"
    ).test(text)
  ) {
    return null;
  }
  const open = TYPES_OPEN_TAG.exec(text);
  if (!open) {
    throw new DocxExportError(
      "malformed-xml",
      `${CONTENT_TYPES_PATH} has no Types element`
    );
  }
  const rootName = /^<([^\s>]+)/.exec(open[0])?.[1] ?? "Types";
  const separator = rootName.indexOf(":");
  const prefix = separator < 0 ? "" : `${rootName.slice(0, separator)}:`;
  const declaration = `<${prefix}Override PartName="${partName}" ContentType="${contentType}"/>`;
  const at = open.index + open[0].length;
  return encodeUtf8(text.slice(0, at) + declaration + text.slice(at), hadBom);
}
