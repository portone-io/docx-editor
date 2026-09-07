/**
 * The contract every part an export writes beside the body is written under.
 *
 * A planner answers with the parts it rewrites, keyed by path, and declares what a part it adds
 * needs through the two writers the context carries: a relationship from the main part and a
 * content type. Those two parts are written once at the end from everything every planner asked
 * for, so no planner writes either of them itself and none can write over what another declared.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { ContentTypeWriter } from "./packageParts";
import type { RelationshipWriter } from "./relationships";
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
