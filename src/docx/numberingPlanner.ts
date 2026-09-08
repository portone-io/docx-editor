/**
 * The numbering part with the definition every list started during editing was started with
 * spliced in, and the part itself when the document arrived without one.
 *
 * The original text is left as is and only the new definitions go in. A document that has no
 * numbering part gets one written from scratch, with the relationship that finds it and the
 * declaration of what it holds asked of the context, so the first list a document takes costs it
 * nothing outside this planner. Nothing is written when no list was started.
 */

import { addListDefinitions } from "../numbering/writeNumbering";
import { elementXml, xmlnsAttr } from "../ooxml/element";
import { NAMESPACES, wName } from "../ooxml/names";
import { ensureRootDeclarations } from "../ooxml/partSplice";
import { decodeUtf8, encodeUtf8 } from "../ooxml/xml";
import {
  NUMBERING_CONTENT_TYPE,
  NUMBERING_REL_TYPE,
  numberingPartOf,
  startedLists,
} from "./newLists";
import { availablePartPath } from "./packageParts";
import type { PartPlanner } from "./partPlan";
import { directoryOf } from "./relationships";

/** A numbering part holding nothing, whose root declares the one prefix everything spliced into it is written under */
function emptyNumberingPart(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    elementXml(wName("numbering"), [xmlnsAttr("w")])
  );
}

export const numberingPlanner: PartPlanner = {
  name: "numbering",
  plan(doc, session, context) {
    // A list no definition was registered for is refused by the invariant list before anything is
    // written, so every list that reaches here has the definition it was started with
    const { defined } = startedLists(doc, session);
    if (defined.size === 0) return null;

    const original = numberingPartOf(session);
    if (original === null) {
      // A dangling relationship already chooses the part's location. Adding a second relationship
      // would leave readers following the first one to an absent definition.
      const path =
        session.numberingPartPath ??
        availablePartPath(session.parts, session.mainPartPath, "numbering");
      if (session.numberingPartPath === null) {
        context.relationships.add({
          type: NUMBERING_REL_TYPE,
          target: path.slice(directoryOf(session.mainPartPath).length),
        });
      }
      context.contentTypes.addOverride(path, NUMBERING_CONTENT_TYPE);
      return new Map([
        [
          path,
          encodeUtf8(addListDefinitions(emptyNumberingPart(), defined), false),
        ],
      ]);
    }

    const { text, hadBom } = decodeUtf8(original.bytes);
    const rewritten = ensureRootDeclarations(
      addListDefinitions(text, defined),
      { namespaces: { w: NAMESPACES.w } }
    );
    return new Map([[original.path, encodeUtf8(rewritten, hadBom)]]);
  },
};
