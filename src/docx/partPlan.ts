/**
 * The contract every part an export writes beside the body is written under.
 *
 * A planner answers with the parts it rewrites, keyed by path, and declares what a part it adds
 * needs through the two writers the context carries: a relationship from the main part and a
 * content type. Those two parts are written once at the end from everything every planner asked
 * for, so no planner writes either of them itself and none can write over what another declared.
 */

import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import { decodeUtf8, parseXml } from "../ooxml/xml";
import { CONTENT_TYPES_PATH, type ContentTypeWriter } from "./packageParts";
import { type RelationshipWriter, relsPathOf } from "./relationships";
import type { SessionStore } from "./session";

export interface PartPlanContext {
  readonly relationships: RelationshipWriter;
  readonly contentTypes: ContentTypeWriter;
}

export interface PartPlanner {
  readonly name: string;
  /** The parts to write, keyed by path, and null when the document gives this planner nothing to write */
  plan(
    doc: PMNode,
    session: SessionStore,
    context: PartPlanContext
  ): ReadonlyMap<string, Uint8Array> | null;
}

/**
 * Every part the planners write, folded into one map in planner order.
 *
 * A planner writing the content types part or the main part's relationships, which the context's
 * writers own, or a part an earlier planner wrote already, is a planner written wrong: either
 * would be written over in silence, so it is refused where it can be seen.
 */
export function runPartPlanners(
  planners: readonly PartPlanner[],
  doc: PMNode,
  session: SessionStore,
  context: PartPlanContext
): Map<string, Uint8Array> {
  const owned = new Set([CONTENT_TYPES_PATH, relsPathOf(session.mainPartPath)]);
  const parts = new Map<string, Uint8Array>();
  for (const planner of planners) {
    for (const [path, bytes] of planner.plan(doc, session, context) ?? []) {
      if (owned.has(path)) {
        throw new Error(
          `the ${planner.name} planner wrote ${path}, which the context's writers own`
        );
      }
      if (parts.has(path)) {
        throw new Error(
          `the ${planner.name} planner wrote ${path}, which an earlier planner wrote already`
        );
      }
      parts.set(path, bytes);
    }
  }
  return parts;
}

const XML_PART = /\.(?:xml|rels)$/i;

/**
 * Every rewritten XML part has to read back as XML, and a refusal names the part.
 *
 * The parts are spliced as text, so this is where a splice that cut in the wrong place shows up:
 * before the package is repacked rather than when somebody opens it.
 */
export function assertPartsParse(
  replacements: ReadonlyMap<string, Uint8Array>
): void {
  for (const [path, bytes] of replacements) {
    if (!XML_PART.test(path)) continue;
    try {
      parseXml(decodeUtf8(bytes).text);
    } catch (cause) {
      throw new DocxExportError(
        "malformed-xml",
        `${path} as written could not be parsed`,
        { cause }
      );
    }
  }
}
