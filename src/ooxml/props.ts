/**
 * Handles the XML fragments that carry formatting, such as `<w:rPr>`, `<w:pPr>`, and `<w:tcPr>`.
 *
 * A fragment is never rebuilt: only the one child a job names is sliced out and swapped, so the
 * rest (borders, shading, margins) stays as the original text wrote it. OOXML lays down the order
 * of the children, so the spot to insert a child that was not there is found by that same order.
 */

import { childOrderOf } from "./childOrder";
import { attrsText, emptyTagXml, type XmlAttr } from "./element";
import { wName } from "./names";
import { parseAttrs, readTag, type Tag } from "./tagScan";
import { childValue } from "./units";
import { elementChildren, localPart, namespaceDecls, parseXml } from "./xml";

export interface PropsChild {
  /** The name with its namespace prefix stripped off (e.g. `gridSpan`) */
  name: string;
  /** This child's original XML fragment exactly as it was */
  xml: string;
  /**
   * Whatever stood between the child before this one and this one: line breaks a producer laid
   * out, comments, anything that is not an element. Carried so that rewriting one child does not
   * quietly drop the rest of what the fragment said. Absent rather than empty when nothing did.
   */
  before?: string;
}

export interface Props {
  /** The opening tag name exactly as written (e.g. `w:tcPr`) */
  tag: string;
  attrs: string | null;
  children: PropsChild[];
  /** The same, for what stood between the last child and the closing tag */
  tail?: string;
}

function rawAttrs(source: string, tag: Tag): string | null {
  const closeLength = tag.kind === "empty" ? 2 : 1;
  const attrs = source.slice(tag.nameEnd, tag.end - closeLength).trim();
  return attrs.length > 0 ? attrs : null;
}

/**
 * What stood inside a single element, and "" for one that stood empty or cannot be made out.
 *
 * The opening tag is read rather than scanned for, so a `>` inside an attribute value does not
 * pass for the end of it.
 */
export function innerXml(xml: string): string {
  const open = readTag(xml, 0);
  if (!open || xml[0] !== "<" || open.kind !== "open") return "";
  const close = xml.lastIndexOf("</");
  return close > open.end ? xml.slice(open.end, close) : "";
}

/**
 * Splits a formatting fragment into its opening tag and its list of children.
 * null if its shape cannot be made out (in which case leaving the original untouched is the safe move).
 */
export function parseProps(xml: string): Props | null {
  const open = readTag(xml, 0);
  if (!open || xml[0] !== "<") return null;
  const attrs = rawAttrs(xml, open);
  if (open.kind === "empty") {
    return open.end === xml.length
      ? { tag: open.name, attrs, children: [] }
      : null;
  }
  if (open.kind !== "open") return null;

  const children: PropsChild[] = [];
  let depth = 0;
  let childStart = -1;
  let childName = "";
  let i = open.end;
  // Everything since the last child ended, which is picked up whole when the next one starts
  let gapStart = open.end;
  const gapBefore = (start: number): { before?: string } => {
    const gap = xml.slice(gapStart, start);
    return gap.length > 0 ? { before: gap } : {};
  };

  while (i < xml.length) {
    const lt = xml.indexOf("<", i);
    if (lt === -1) return null;
    const tag = readTag(xml, lt);
    if (!tag) return null;

    if (tag.kind === "other") {
      i = tag.end;
      continue;
    }

    if (depth === 0) {
      if (tag.kind === "close") {
        if (tag.name !== open.name || tag.end !== xml.length) return null;
        const tail = xml.slice(gapStart, lt);
        return tail.length > 0
          ? { tag: open.name, attrs, children, tail }
          : { tag: open.name, attrs, children };
      }
      childStart = lt;
      childName = tag.name;
      if (tag.kind === "empty") {
        children.push({
          name: localPart(childName),
          xml: xml.slice(childStart, tag.end),
          ...gapBefore(childStart),
        });
        gapStart = tag.end;
      } else {
        depth = 1;
      }
    } else if (tag.kind === "open") {
      depth += 1;
    } else if (tag.kind === "close") {
      depth -= 1;
      if (depth === 0) {
        children.push({
          name: localPart(childName),
          xml: xml.slice(childStart, tag.end),
          ...gapBefore(childStart),
        });
        gapStart = tag.end;
      }
    }
    i = tag.end;
  }
  return null;
}

/**
 * With nothing to say, the formatting fragment itself is not written. A fragment holding only
 * whitespace says nothing; one holding a comment does, so it is written back.
 */
export function renderProps(props: Props): string {
  const tail = props.tail ?? "";
  if (props.children.length === 0 && tail.trim().length === 0) return "";
  const open = props.attrs ? `<${props.tag} ${props.attrs}>` : `<${props.tag}>`;
  const inner = props.children
    .map((child) => (child.before ?? "") + child.xml)
    .join("");
  return open + inner + tail + `</${props.tag}>`;
}

/**
 * The spot this child goes into within the prescribed order: ahead of the first child the order
 * puts after it. A child already there that the order does not know decides nothing, so a new
 * child lands behind it when it stands ahead of the spot and ahead of it otherwise.
 *
 * Throws for a name the order does not know at all. The order mirrors the schema for its parent,
 * so such a name is either an entry the registry is missing or a child the parent may not hold,
 * and writing it on the end would hide either one behind an export a validator may refuse.
 */
function insertIndex(
  children: readonly PropsChild[],
  name: string,
  order: readonly string[],
  tag: string
): number {
  const target = order.indexOf(name);
  if (target === -1) {
    throw new Error(
      `${name} is not a child the order of ${tag} knows; add it to CHILD_ORDER`
    );
  }
  const at = children.findIndex((child) => order.indexOf(child.name) > target);
  return at === -1 ? children.length : at;
}

/**
 * Drops every child of this name, carrying what stood in front of each onto whatever follows.
 * A gap belongs to the fragment rather than to the child it happened to sit in front of, so
 * removing a child does not take a producer's comment away with it.
 */
function dropChildren(
  children: readonly PropsChild[],
  name: string
): { kept: PropsChild[]; carried: string } {
  const kept: PropsChild[] = [];
  let carried = "";
  for (const child of children) {
    const before = carried + (child.before ?? "");
    if (child.name === name) {
      carried = before;
      continue;
    }
    carried = "";
    kept.push(
      before === ""
        ? { name: child.name, xml: child.xml }
        : { ...child, before }
    );
  }
  return { kept, carried };
}

/** What a dropped child left in front of it, with nothing left to follow, belongs to the tail */
function tailWith(props: Props, carried: string): { tail?: string } {
  const tail = carried + (props.tail ?? "");
  return tail === "" ? {} : { tail };
}

/**
 * Replaces a single child with new XML.
 *
 * A null `xml` removes that child. A child that was not there goes into the spot `CHILD_ORDER`
 * lays down for this fragment's own tag, so the caller names what it is writing and not where it
 * goes. Naming the element this fragment stands inside picks the right order where one element
 * carries children in a different order under a different parent (`pPr/rPr`).
 */
export function setChild(
  props: Props,
  name: string,
  xml: string | null,
  parent?: string
): Props {
  return setChildren(props, name, xml === null ? [] : [xml], parent);
}

/**
 * Replaces every child of one name with this list of them, in the order given.
 *
 * `CT_SectPr` lets a section name a header and a footer once per variant, so a name standing more
 * than once is a list rather than a single child, and writing one of them is writing the list. An
 * empty list takes them all away.
 */
export function setChildren(
  props: Props,
  name: string,
  xmls: readonly string[],
  parent?: string
): Props {
  const order = childOrderOf(localPart(props.tag), parent);
  const at = props.children.findIndex((entry) => entry.name === name);

  // The first one keeps the spot it originally occupied, and what stood in front of it; the rest
  // follow it, and any further spot the same name held is given up
  if (at !== -1 && xmls.length > 0) {
    const rest = dropChildren(props.children.slice(at + 1), name);
    return {
      ...props,
      children: [
        ...props.children.slice(0, at),
        ...xmls.map((xml, index) =>
          index === 0 ? { ...props.children[at], xml } : { name, xml }
        ),
        ...rest.kept,
      ],
      ...tailWith(props, rest.carried),
    };
  }

  const { kept, carried } = dropChildren(props.children, name);
  if (xmls.length === 0) {
    return { ...props, children: kept, ...tailWith(props, carried) };
  }
  const index = insertIndex(kept, name, order, props.tag);
  return {
    ...props,
    children: [
      ...kept.slice(0, index),
      ...xmls.map((xml) => ({ name, xml })),
      ...kept.slice(index),
    ],
    ...tailWith(props, carried),
  };
}

export function propsChild(
  children: readonly PropsChild[],
  name: string
): PropsChild | undefined {
  return children.find((child) => child.name === name);
}

/**
 * The opening tag's attributes as pairs, and null when their shape cannot be made out.
 *
 * The pairs are worked out on demand rather than kept on `Props`, because an untouched fragment
 * has to go back out with the spacing, the quoting and the escaping its producer chose. Only a
 * fragment something writes to passes through `withAttrs`, which is where they are written again.
 */
export function attrsOf(props: Props): XmlAttr[] | null {
  return props.attrs === null ? [] : parseAttrs(props.attrs);
}

/** The fragment with these attributes written in place of the ones its opening tag carried */
export function withAttrs(props: Props, attrs: readonly XmlAttr[]): Props {
  return { ...props, attrs: attrs.length === 0 ? null : attrsText(attrs) };
}

/**
 * One child of a fragment as it was written: the tag, prefix included, so an edited child keeps
 * the spelling the document chose, and its attributes. A child that is not there has neither, so
 * no tag with attributes hanging off it can be made.
 */
export type ChildElement =
  | { tag: string; attrs: readonly XmlAttr[] }
  | { tag: null; attrs: readonly [] };

/**
 * The tag and attributes of one child. A child that is not there reads as no tag and no
 * attributes, and null says its shape cannot be made out, which leaves the caller to back out
 * rather than write over markup it could not read.
 */
export function childElement(props: Props, name: string): ChildElement | null {
  const child = propsChild(props.children, name);
  if (child === undefined) return { tag: null, attrs: [] };
  const parsed = parseProps(child.xml);
  const attrs = parsed === null ? null : attrsOf(parsed);
  return parsed === null || attrs === null ? null : { tag: parsed.tag, attrs };
}

/**
 * Always writes the element, closing it on its own when it holds nothing. Whitespace alone counts
 * as nothing, the same as `renderProps` reads it, so a pretty-printed fragment emptied of its
 * children collapses rather than keeping the line breaks that stood between them.
 */
export function renderElement(props: Props): string {
  return renderProps(props) || emptyTagXml(props.tag, props.attrs);
}

/**
 * An element whose children go in the order `CHILD_ORDER` lays down, whatever order they are
 * handed in. This is how a fragment written from scratch follows the same order an edited one
 * is held to, and a child the order does not know is refused the same way.
 */
export function orderedElement(
  tag: string,
  attrs: readonly XmlAttr[],
  children: readonly PropsChild[]
): string {
  const empty = withAttrs({ tag, attrs: null, children: [] }, attrs);
  return renderElement(
    children.reduce(
      (props, child) => setChild(props, child.name, child.xml),
      empty
    )
  );
}

function editPath(
  props: Props,
  path: readonly string[],
  edit: (child: Props | null) => Props | null,
  parent: string | undefined
): Props | null {
  const [name, ...rest] = path;
  const current = propsChild(props.children, name);
  const parsed = current === undefined ? null : parseProps(current.xml);
  if (current !== undefined && parsed === null) return null;

  if (rest.length === 0) {
    const next = edit(parsed);
    return setChild(props, name, next && renderElement(next), parent);
  }
  const nested = editPath(
    parsed ?? { tag: wName(name), attrs: null, children: [] },
    rest,
    edit,
    localPart(props.tag)
  );
  if (nested === null) return null;
  const rendered = renderProps(nested);
  return setChild(props, name, rendered === "" ? null : rendered, parent);
}

/**
 * Edits one child down a path of names, parsing only the fragments the path runs through and
 * leaving every other child as the original text wrote it.
 *
 * The callback is handed the child the path names, null for one that is not there, and answers
 * with what it is to become, null to take it away. A container along the path that was not there
 * is written, and one the edit leaves empty is taken away with it. null when a fragment on the
 * path cannot be made out, which leaves the caller to back out.
 */
export function editChild(
  props: Props,
  path: readonly [string, ...string[]],
  edit: (child: Props | null) => Props | null
): Props | null {
  return editPath(props, path, edit, undefined);
}

/**
 * Reads a single formatting fragment into an element.
 *
 * A fragment carries no namespace declarations, so they are put back on as it is wrapped.
 * null if its shape cannot be made out, in which case the caller leaves the display values alone.
 */
export function parsePropsXml(xml: string): Element | null {
  try {
    const wrapped = `<props ${namespaceDecls(xml)}>${xml}</props>`;
    return elementChildren(parseXml(wrapped).documentElement)[0] ?? null;
  } catch {
    return null;
  }
}

/**
 * The style ids already read out of a formatting fragment.
 *
 * The toolbar decides the style of every selected paragraph on each render, and parsing the same
 * fragment over and over is the whole cost of that. The fragment text is the key, so one fragment
 * is parsed once however many paragraphs share it.
 */
const styleIdsByPPr = new Map<string, string | null>();

/** Past this many fragments the table is dropped rather than grown for as long as the page lives */
const STYLE_ID_CACHE_LIMIT = 2000;

/** The style name the paragraph formatting XML points at. null if it points at none */
export function styleIdOf(pPr: unknown): string | null {
  // Paragraphs that point at a style are rare, so we screen them out first with a cheap string check
  if (typeof pPr !== "string" || !pPr.includes("pStyle")) return null;
  const known = styleIdsByPPr.get(pPr);
  if (known !== undefined) return known;
  const el = parsePropsXml(pPr);
  const id = el ? childValue(el, "pStyle") : null;
  if (styleIdsByPPr.size >= STYLE_ID_CACHE_LIMIT) styleIdsByPPr.clear();
  styleIdsByPPr.set(pPr, id);
  return id;
}
