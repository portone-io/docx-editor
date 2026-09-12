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
import type { FidelityCollector } from "./fidelity";
import {
  CONTENT_TYPES_PATH,
  type ContentTypeWriter,
  declaredXmlParts,
} from "./packageParts";
import { type RelationshipWriter, relsPathOf } from "./relationships";
import type { SessionStore } from "./session";

export interface PartPlanContext {
  readonly relationships: RelationshipWriter;
  readonly contentTypes: ContentTypeWriter;
  /** Where a story writer records an approximation it had to make, beside the body writer's */
  readonly notes: FidelityCollector;
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
 * The body and context-owned parts are reserved. Earlier writes include media planned before
 * the body, so no later planner can overwrite those bytes either. Part names are compared
 * without regard to case, as package part names are.
 */
export function runPartPlanners(
  planners: readonly PartPlanner[],
  doc: PMNode,
  session: SessionStore,
  context: PartPlanContext,
  prior: ReadonlyMap<string, Uint8Array> = new Map()
): Map<string, Uint8Array> {
  const owned = new Set(
    [
      CONTENT_TYPES_PATH,
      relsPathOf(session.mainPartPath),
      session.mainPartPath,
    ].map((path) => path.toLowerCase())
  );
  const parts = new Map(prior);
  const written = new Set(
    Array.from(prior.keys(), (path) => path.toLowerCase())
  );
  for (const planner of planners) {
    for (const [path, bytes] of planner.plan(doc, session, context) ?? []) {
      if (owned.has(path.toLowerCase())) {
        throw new Error(
          `the ${planner.name} planner wrote ${path}, which another export writer owns`
        );
      }
      if (written.has(path.toLowerCase())) {
        throw new Error(
          `the ${planner.name} planner wrote ${path}, which an earlier writer wrote already`
        );
      }
      parts.set(path, bytes);
      written.add(path.toLowerCase());
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
  replacements: ReadonlyMap<string, Uint8Array>,
  contentTypes?: Uint8Array
): void {
  const paths = [...replacements.keys()];
  const declared =
    contentTypes && paths.some((path) => !XML_PART.test(path))
      ? declaredXmlParts(contentTypes, paths)
      : new Set<string>();
  for (const [path, bytes] of replacements) {
    if (!XML_PART.test(path) && !declared.has(path)) continue;
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
