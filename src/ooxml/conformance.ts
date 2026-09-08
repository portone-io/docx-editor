/**
 * Which conformance class a package is written in, and whether its main part is one this editor
 * can write back into. Both are settled once, when the file is opened.
 *
 * ECMA-376 defines two document conformance classes, Strict (Part 1 §2.1) and Transitional
 * (Part 4 §2.1). They describe the same vocabulary under different names: the main document part's
 * root namespace and the relationship the package reaches it through differ between the two
 * (Part 1 §11.3.10, Part 4 §9.2.10). This editor supports a subset of Transitional. Strict is
 * refused explicitly so that its markup is never read using Transitional assumptions.
 *
 * A namespace prefix is the producer's own choice and carries no meaning of its own, so readers
 * identify markup by namespace URI and local name. The WordprocessingML writer emits `w`
 * (`ooxml/names`), so the main part's root must bind that prefix to the Transitional namespace.
 */

import { W_NS, W_PREFIX } from "./names";

/** The WordprocessingML namespace of a Strict main document part (Part 1 §11.3.10) */
export const STRICT_W_NS = "http://purl.oclc.org/ooxml/wordprocessingml/main";

/** The relationship a Strict package reaches its main document part through (Part 1 §11.3.10) */
export const STRICT_OFFICE_DOCUMENT_REL =
  "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";

export type Conformance = "transitional" | "strict";

/** The class this main part root is written in, null for a root of neither vocabulary */
export function conformanceOf(root: Element): Conformance | null {
  if (root.namespaceURI === W_NS) return "transitional";
  if (root.namespaceURI === STRICT_W_NS) return "strict";
  return null;
}

/**
 * Whether the root binds the prefix every writer emits (`w`) to the WordprocessingML namespace.
 *
 * A root has nothing above it to inherit a declaration from, so what it writes itself is the whole
 * answer. A default namespace or another prefix alone does not bind `w`.
 */
export function bindsWritingPrefix(root: Element): boolean {
  return Array.from(root.attributes).some(
    (attr) => attr.name === `xmlns:${W_PREFIX}` && attr.value === W_NS
  );
}
