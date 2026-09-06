/**
 * The markup-compatibility preprocessing of ECMA-376 part 3, which every consumer performs
 * before reading a part against the part 1 schemas.
 *
 * The part 1 schemas describe none of it: a namespace a part declares ignorable may hold
 * attributes and elements the schema has no `anyAttribute` or `xsd:any` to admit, and
 * `mc:AlternateContent` is not in the wordprocessing vocabulary at all. Reading a part without
 * this step therefore fails a document a conforming consumer accepts.
 *
 * What is implemented is the ignorable-namespace and `AlternateContent` handling that documents
 * this package reads and writes rely on. `mc:ProcessContent`, `mc:PreserveElements` and
 * `mc:PreserveAttributes` are not: no part the tests validate carries them, and a part that does
 * would keep markup here that a conforming consumer had dropped.
 */

import { parseXml } from "../../ooxml/xml";

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const NOTHING_IGNORABLE: ReadonlySet<string> = new Set();

/**
 * The namespaces ignorable at this element: the ones its ancestors declared, plus the prefixes
 * of its own `mc:Ignorable` resolved where they are written. A prefix naming no namespace in
 * scope is left out, since there is no markup for it to name.
 */
function ignorableAt(
  el: Element,
  inherited: ReadonlySet<string>
): ReadonlySet<string> {
  const declared = el.getAttributeNS(MC_NS, "Ignorable");
  if (declared === null) return inherited;
  const namespaces = new Set(inherited);
  for (const prefix of declared.split(/\s+/)) {
    const namespace = prefix === "" ? null : el.lookupNamespaceURI(prefix);
    if (namespace !== null) namespaces.add(namespace);
  }
  return namespaces;
}

function isIgnorable(el: Element, ignorable: ReadonlySet<string>): boolean {
  return el.namespaceURI !== null && ignorable.has(el.namespaceURI);
}

function childByName(parent: Element, name: string): Element | null {
  return (
    Array.from(parent.children).find(
      (child) => child.namespaceURI === MC_NS && child.localName === name
    ) ?? null
  );
}

/** The compatibility attributes and the ones an ignorable namespace holds */
function stripAttributes(el: Element, ignorable: ReadonlySet<string>): void {
  for (const attribute of Array.from(el.attributes)) {
    const namespace = attribute.namespaceURI;
    if (namespace === null) continue;
    if (namespace === MC_NS || ignorable.has(namespace)) {
      el.removeAttributeNS(namespace, attribute.localName);
    }
  }
}

function rewriteInside(el: Element, inherited: ReadonlySet<string>): void {
  const ignorable = ignorableAt(el, inherited);
  for (const child of Array.from(el.children)) {
    rewriteChild(child, el, ignorable);
  }
  stripAttributes(el, ignorable);
}

function rewriteChild(
  el: Element,
  parent: Element,
  ignorable: ReadonlySet<string>
): void {
  // The consumer that understands none of the alternatives takes the fallback content, which
  // stands where the `mc:AlternateContent` stood and is processed there in its turn
  if (el.namespaceURI === MC_NS && el.localName === "AlternateContent") {
    const fallback = childByName(el, "Fallback");
    for (const kept of fallback === null ? [] : Array.from(fallback.children)) {
      parent.insertBefore(kept, el);
      rewriteChild(kept, parent, ignorable);
    }
    el.remove();
    return;
  }
  if (isIgnorable(el, ignorable)) {
    el.remove();
    return;
  }
  rewriteInside(el, ignorable);
}

/**
 * The part as a consumer of the part 1 vocabulary reads it.
 *
 * Test-only, so the document goes back out through the DOM rather than being spliced as text:
 * the shape of the result is what a validator reads, not what a writer emits.
 */
export function withoutIgnorableMarkup(xml: string): string {
  const doc = parseXml(xml);
  rewriteInside(doc.documentElement, NOTHING_IGNORABLE);
  return new XMLSerializer().serializeToString(doc);
}
