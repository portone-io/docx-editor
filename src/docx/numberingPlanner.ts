/**
 * The numbering part with the definition of every list started during editing spliced in.
 *
 * The original text is left as is and only the new definitions go in. Nothing is written when no
 * list was started.
 */

import { addListDefinitions } from "../numbering/writeNumbering";
import { NAMESPACES } from "../ooxml/names";
import { ensureRootDeclarations } from "../ooxml/partSplice";
import { decodeUtf8, encodeUtf8 } from "../ooxml/xml";
import { newNumIds, numberingPartOf } from "./newLists";
import type { PartPlanner } from "./partPlan";

export const numberingPlanner: PartPlanner = {
  name: "numbering",
  plan(doc, session) {
    const added = newNumIds(doc, session);
    const original = numberingPartOf(session);
    // A new list in a document with no numbering.xml is refused by the invariant list before
    // anything is written, so a missing part here has nothing to hold
    if (added.length === 0 || original === null) return null;

    const { text, hadBom } = decodeUtf8(original.bytes);
    const rewritten = ensureRootDeclarations(addListDefinitions(text, added), {
      namespaces: { w: NAMESPACES.w },
    });
    return new Map([[original.path, encodeUtf8(rewritten, hadBom)]]);
  },
};
