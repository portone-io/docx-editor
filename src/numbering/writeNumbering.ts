/**
 * Splices the definitions the newly started lists were registered with into the original
 * numbering.xml.
 *
 * Not one character of the original is altered; new elements are merely slotted in.
 * OOXML requires the definitions (`abstractNum`) to come before the numbers (`num`), and
 * `CHILD_ORDER.numbering` is what places each where it belongs among what the part holds.
 */

import { splicePart } from "../ooxml/partSplice";
import { ST_DecimalNumber } from "../ooxml/simpleTypes";
import { wAttr } from "../ooxml/units";
import { elementChildren, parseXml } from "../ooxml/xml";
import { abstractNumXml, numberingIdAllocator, numXml } from "./listTemplate";
import type { NewList } from "./parseNumbering";

/** The definition ids a new definition must not reuse. */
function abstractNumIds(xml: string): number[] {
  return elementChildren(parseXml(xml).documentElement).flatMap((child) => {
    if (child.localName !== "abstractNum") return [];
    const id = ST_DecimalNumber.parse(wAttr(child, "abstractNumId"));
    return id === null ? [] : [id];
  });
}

/**
 * numbering.xml with the definition each of these lists was started with added under its number.
 * When there is nothing to add, the original string is returned unchanged.
 */
export function addListDefinitions(
  xml: string,
  lists: ReadonlyMap<number, NewList>
): string {
  if (lists.size === 0) return xml;

  const takeId = numberingIdAllocator(abstractNumIds(xml));
  const additions = [...lists]
    .sort(([left], [right]) => left - right)
    .map(([numId, list]) => ({
      numId,
      list,
      abstractNumId: takeId(),
    }));

  return splicePart(xml, {
    root: "numbering",
    insert: [
      ...additions.map((added) => ({
        name: "abstractNum",
        xml: abstractNumXml(added.abstractNumId, added.list),
      })),
      ...additions.map((added) => ({
        name: "num",
        xml: numXml(added.numId, added.abstractNumId),
      })),
    ],
  });
}
