/**
 * The gate a raw OOXML string passes through on its way from the DOM into a node or mark attr.
 *
 * A preserved fragment goes back out the way it came, spliced into a slot the writer opens and
 * closes around it (`docx/serializeParagraph`), so a fragment that closes that slot itself, opens
 * a sibling beside it, or lays text down beside what it holds writes something the editor never
 * modelled into the exported file. What stands inside the element it carries is that element's own
 * and travels with it. Reading one back is the only place such a string can arrive from outside,
 * so every attr that carries one is read through here and a fragment that does not hold its shape
 * is turned down rather than corrected.
 *
 * A refusal is not an error: `getAttrs` answers `false` with it, which drops the parse rule and
 * leaves ProseMirror to settle the content one level plainer. Demotion, not contamination.
 */

import {
  declaredNamespaces,
  elementChildren,
  namespaceDecls,
  parseXml,
  RESERVED_PREFIXES,
  W_NS,
} from "./xml";

/**
 * What a raw XML string has to look like to be let into the attr that carries it.
 *
 * `element` is a whole element: a properties fragment, a drawing, an annotation reference.
 * `attributes` is what stood inside an opening tag, as `attrString` writes it.
 * `openTag` is an opening tag with everything it wrapped cut away, which the writer puts back by
 * appending the closing text `closedBy` names. `head` is what may still stand between the two:
 * the property elements the tag carries ahead of its content, by local name.
 */
export type RawXmlShape =
  | { kind: "element"; names: readonly string[] | "any" }
  | { kind: "attributes" }
  | {
      kind: "openTag";
      name: string;
      closedBy: string;
      head: readonly string[];
    };

type OpenTagShape = Extract<RawXmlShape, { kind: "openTag" }>;

/** A single WordprocessingML element going by one of these local names */
export function ELEMENT(...names: string[]): RawXmlShape {
  return { kind: "element", names };
}

/**
 * A single element of any name, in any namespace.
 *
 * What import could not model keeps its own XML (`rawInline`, `rawBlock`), and that is a
 * `m:oMathPara`, a `mc:AlternateContent`, a `w:ins`, a bookmark: naming the ones allowed would
 * turn a document the editor reads today into a demoted one.
 */
export const ANY_ELEMENT: RawXmlShape = { kind: "element", names: "any" };

/** The attributes of an opening tag, with neither the tag nor anything it held */
export const ATTRIBUTES: RawXmlShape = { kind: "attributes" };

/**
 * The one element this fragment holds. null when it holds anything else: nothing, more than one
 * element, or a sibling of any other kind beside it.
 *
 * The wrapper is what makes a fragment parseable at all, since it carries no namespace
 * declarations of its own, and it is also what catches a fragment that closes its own parent:
 * that one no longer nests inside the wrapper and does not parse.
 */
function loneElement(xml: string): Element | null {
  let root: Element;
  try {
    root = parseXml(`<x ${namespaceDecls(xml)}>${xml}</x>`).documentElement;
  } catch {
    return null;
  }
  if (root.childNodes.length !== 1) return null;
  return elementChildren(root)[0] ?? null;
}

function isNamed(el: Element, names: readonly string[]): boolean {
  return el.namespaceURI === W_NS && names.includes(el.localName);
}

/**
 * Whether the opened tag holds its own head and nothing besides what `closedBy` brought.
 *
 * The writer splices the content it modelled between the opening tag and the closing text, so a
 * fragment that carried children of its own would put them in front of that content, and one that
 * opened the very slot the closing text opens would leave the file with two of them. Only the
 * property elements `head` names may stand there, and no text at all.
 */
function holdsHeadAlone(el: Element, shape: OpenTagShape): boolean {
  const closing = loneElement(`<w:${shape.name}>${shape.closedBy}`);
  if (closing === null) return false;
  const children = elementChildren(el);
  const own = children.slice(
    0,
    children.length - elementChildren(closing).length
  );
  return (
    el.childNodes.length === children.length &&
    own.every((child) => isNamed(child, shape.head))
  );
}

/**
 * Whether the fragment binds a prefix the writer depends on to a namespace of its own choosing.
 *
 * A declaration travels with the tag it stands on and covers everything under it, and for an
 * attribute list or an opening tag that is the very subtree the writer splices its own markup into:
 * `xmlns:r="urn:evil"` on a paragraph leaves every `r:id` and `r:embed` written under it naming a
 * relationship the package does not have, so the link re-imports with no address and the image
 * loses its part. A declaration that agrees with what the prefix already means, and one binding a
 * prefix the editor reads nothing under, are the producer's own and travel untouched.
 */
function rebindsReservedPrefix(xml: string): boolean {
  return Array.from(declaredNamespaces(xml)).some(([prefix, uri]) => {
    const reserved = RESERVED_PREFIXES.get(prefix);
    return reserved !== undefined && reserved !== uri;
  });
}

function holdsShape(shape: RawXmlShape, value: string): boolean {
  if (rebindsReservedPrefix(value)) return false;
  switch (shape.kind) {
    case "element": {
      const el = loneElement(value);
      if (el === null) return false;
      return shape.names === "any" || isNamed(el, shape.names);
    }
    case "attributes": {
      // The list is read on a tag of its own rather than beside the declarations that make it
      // parseable, so that a namespace the list declares shadows them instead of standing twice on
      // one element, which is what a producer writing `<w:p xmlns:w14="..." w14:paraId="...">`
      // hands us
      const el = loneElement(`<y ${value}></y>`);
      // An attribute list that opened a child of its own ended the tag it was written into
      return el !== null && el.childNodes.length === 0;
    }
    case "openTag": {
      // Which element the fragment opened is settled by parsing at all: the closing text ends with
      // that element's own end tag, and XML matches an end tag against the very spelling the start
      // tag used, so nothing but `<w:{name}` can reach here
      const el = loneElement(value + shape.closedBy);
      return el !== null && holdsHeadAlone(el, shape);
    }
  }
}

/**
 * The value if it holds the shape, null if there was none, `false` if it does not.
 *
 * The three answers are what `getAttrs` needs: a value to carry, an absent attr, and a rule to
 * give up on.
 */
export function acceptRawXml(
  shape: RawXmlShape,
  value: string | null
): string | null | false {
  if (value === null) return null;
  return holdsShape(shape, value) ? value : false;
}
