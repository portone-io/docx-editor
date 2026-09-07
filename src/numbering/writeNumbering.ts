/**
 * Splices the definitions of newly started lists into the original numbering.xml.
 *
 * Not one character of the original is altered; new elements are merely slotted in.
 * OOXML requires the definitions (`abstractNum`) to come before the numbers (`num`), and
 * `CHILD_ORDER.numbering` is what places each where it belongs among what the part holds.
 */

import { splicePart } from "../ooxml/partSplice";
import { wAttr } from "../ooxml/units";
import { elementChildren, parseXml } from "../ooxml/xml";
import { abstractNumXml, listKindOf, numXml } from "./listTemplate";

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

  const firstAbstractNumId = maxAbstractNumId(xml) + 1;
  const additions = [...numIds]
    .sort((a, b) => a - b)
    .map((numId, index) => ({
      numId,
      abstractNumId: firstAbstractNumId + index,
    }));

  return splicePart(xml, {
    root: "numbering",
    insert: [
      ...additions.map((added) => ({
        name: "abstractNum",
        xml: abstractNumXml(added.abstractNumId, listKindOf(added.numId)),
      })),
      ...additions.map((added) => ({
        name: "num",
        xml: numXml(added.numId, added.abstractNumId),
      })),
    ],
  });
}
