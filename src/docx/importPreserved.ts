/**
 * One element kept exactly as it came, at whatever level of the document it stood.
 *
 * `./importPolicy` says what an element is worth keeping as; this is where that answer becomes a
 * node. Every level has a node of its own - inside a run, beside the runs of a paragraph, and in
 * the place of a block - so the three readers build them the same way and cannot drift apart in
 * what a chip shows or which fragments a guard answers for.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import { serializeXml } from "../ooxml/xml";
import { docxSchema } from "../schema";
import {
  type DemotionPolicy,
  defaultPreservation,
  type ElementPolicy,
  type PreservationRule,
  type PreservingLevel,
  policyFor,
} from "./importPolicy";

/**
 * What one preserved fragment says about itself, read off its rule and its own content.
 *
 * A chip shows the text the element held, which is the field result a reader sees and the words a
 * tracked insertion put there, so the chip says what it stands for rather than only that it is one.
 */
function preservedText(el: Element, rule: PreservationRule): string | null {
  if (rule.display === "text") return rule.text ?? null;
  if (rule.display !== "chip") return null;
  return el.textContent === "" ? null : (el.textContent ?? null);
}

function preservedAttrs(
  el: Element,
  rule: PreservationRule
): Record<string, unknown> {
  return {
    xml: serializeXml(el),
    element: el.localName,
    display: rule.display,
    text: preservedText(el, rule),
    guarded: rule.guarded,
  };
}

/**
 * The rule an element is kept under.
 *
 * The level's own narrowest preservation stands in wherever the table said `model` and the reader
 * could not read the element after all, so a structure it could not take apart costs that element
 * and no more.
 */
export function preservationOf(
  policy: ElementPolicy | DemotionPolicy,
  level: PreservingLevel
): PreservationRule {
  return policy.tier === "model" || policy.tier === "demote"
    ? defaultPreservation(level)
    : policy;
}

/** One child of a run kept as it came, wearing the run mark so it goes back inside that run */
export function buildPreservedRunContent(
  el: Element,
  marks: readonly Mark[],
  rule: PreservationRule
): PMNode {
  return docxSchema.nodes.rawRunContent.create(
    preservedAttrs(el, rule),
    null,
    marks
  );
}

/** One child of a paragraph, or of a wrapper inside it, kept as it came */
export function buildPreservedInline(
  el: Element,
  wrappers: readonly Mark[],
  rule: PreservationRule
): PMNode {
  return docxSchema.nodes.rawInline.create(
    preservedAttrs(el, rule),
    null,
    wrappers
  );
}

/**
 * One block kept as it came, standing where it stood.
 *
 * A body block is one of the fragments the session holds, so it names that fragment and the bytes
 * it arrived as go back out untouched. A block inside a cell was never a fragment of its own, so
 * it carries its XML along with it.
 */
export function buildPreservedBlock(
  el: Element,
  srcId: string | null,
  level: "body" | "tc"
): PMNode {
  const rule = preservationOf(policyFor(el, level), level);
  return docxSchema.nodes.rawBlock.create({
    xml: srcId === null ? serializeXml(el) : null,
    srcId,
    name: el.nodeName,
    display: rule.display,
    guarded: rule.guarded,
  });
}
