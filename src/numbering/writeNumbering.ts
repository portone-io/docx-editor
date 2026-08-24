/**
 * Splices the definitions of newly started lists into the original numbering.xml.
 *
 * Not one character of the original is altered; new elements are merely slotted in at
 * two places.
 * OOXML requires the definitions (`abstractNum`) to come before the numbers (`num`), so
 * the definitions go in front of the first number and the numbers go at the very end.
 */

import { DocxExportError } from "../ooxml/errors";
import { wAttr } from "../ooxml/units";
import { elementChildren, parseXml } from "../ooxml/xml";
import { abstractNumXml, listKindOf, numXml } from "./listTemplate";

/**
 * The first position where an opening tag with that name appears, whatever its namespace
 * prefix
 */
function openTagAt(xml: string, name: string): number | null {
  const match = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${name}[\\s/>]`).exec(xml);
  return match ? match.index : null;
}

/**
 * The position that closes the root element. There is only one in the document, so the
 * last match is used
 */
function closeTagAt(xml: string, name: string): number | null {
  const matches = Array.from(
    xml.matchAll(new RegExp(`</(?:[A-Za-z_][\\w.-]*:)?${name}\\s*>`, "g"))
  );
  return matches.at(-1)?.index ?? null;
}

/** The largest definition id already in use */
function maxAbstractNumId(xml: string): number {
  let max = 0;
  for (const child of elementChildren(parseXml(xml).documentElement)) {
    if (child.localName !== "abstractNum") continue;
    const id = Number.parseInt(wAttr(child, "abstractNumId") ?? "", 10);
    if (Number.isFinite(id)) max = Math.max(max, id);
  }
  return max;
}

/**
 * numbering.xml with a standard template definition added for every list number.
 * When there is nothing to add, the original string is returned unchanged.
 */
export function addListDefinitions(
  xml: string,
  numIds: readonly number[]
): string {
  if (numIds.length === 0) return xml;

  const rootEnd = closeTagAt(xml, "numbering");
  if (rootEnd === null) {
    throw new DocxExportError(
      "malformed-xml",
      "no closing tag found in numbering.xml"
    );
  }
  // If an element that has to come last is present, the new numbers go in front of it
  const numsAt = openTagAt(xml, "numIdMacAtCleanup") ?? rootEnd;
  const definitionsAt = openTagAt(xml, "num") ?? numsAt;

  const firstAbstractNumId = maxAbstractNumId(xml) + 1;
  const additions = [...numIds]
    .sort((a, b) => a - b)
    .map((numId, index) => ({
      numId,
      abstractNumId: firstAbstractNumId + index,
    }));

  const definitions = additions
    .map((added) =>
      abstractNumXml(added.abstractNumId, listKindOf(added.numId))
    )
    .join("");
  const nums = additions
    .map((added) => numXml(added.numId, added.abstractNumId))
    .join("");

  return (
    xml.slice(0, definitionsAt) +
    definitions +
    xml.slice(definitionsAt, numsAt) +
    nums +
    xml.slice(numsAt)
  );
}
