/**
 * Writing one element out as text, and merging the attributes of one.
 *
 * Everything here works on strings, because that is what the writers hold: a fragment the file
 * arrived with is carried around as its original text, and an edit swaps out one element or one
 * attribute of it. Rebuilding a parsed element instead would rewrite markup nobody touched.
 *
 * An attribute is matched by its local part rather than by its written name, so a document that
 * bound the WordprocessingML namespace to a prefix of its own is edited in the spelling it chose.
 */

import { type KnownPrefix, NAMESPACES, wName } from "./names";
import { escapeXml, localPart } from "./xml";

/** One attribute as it is written: the whole name, prefix included, and the value before escaping */
export type XmlAttr = readonly [name: string, value: string];

/** `a="b" c="d"`. Empty for no attributes, so an opening tag needs no separator of its own */
export function attrsText(attrs: readonly XmlAttr[]): string {
  return attrs
    .map(([name, value]) => `${name}="${escapeXml(value)}"`)
    .join(" ");
}

/**
 * `<name a="b"/>`, or `<name a="b">children</name>` for an element that holds something.
 *
 * Attribute values are escaped here; children are spliced in as they are, being markup already, so
 * a caller writing text as a child escapes it itself. An element handed no child at all closes on
 * its own, which means a caller that wants an empty pair of tags passes the empty text as a child.
 */
export function elementXml(
  name: string,
  attrs: readonly XmlAttr[],
  children: readonly string[] = []
): string {
  const text = attrsText(attrs);
  const open = text === "" ? `<${name}` : `<${name} ${text}`;
  if (children.length === 0) return `${open}/>`;
  return `${open}>${children.join("")}</${name}>`;
}

/**
 * The namespace declaration a fragment carries when the part it is spliced into does not declare
 * the prefix itself.
 */
export function xmlnsAttr(prefix: KnownPrefix): XmlAttr {
  return [`xmlns:${prefix}`, NAMESPACES[prefix]];
}

/**
 * The attributes with the one of this local name written as `value`, or removed when it is null.
 *
 * One already there keeps the slot it sat in and the name it was written under, so an edited
 * element reads as the original did. One that is not there yet goes on the end under the `w`
 * prefix.
 */
export function withAttr(
  attrs: readonly XmlAttr[],
  name: string,
  value: string | null
): XmlAttr[] {
  if (value === null) return withoutAttrs(attrs, [name]);
  const at = attrs.findIndex(([attr]) => localPart(attr) === name);
  if (at === -1) return [...attrs, [wName(name), value]];
  return attrs.map(
    (attr, index): XmlAttr => (index === at ? [attr[0], value] : attr)
  );
}

export function withoutAttrs(
  attrs: readonly XmlAttr[],
  localNames: readonly string[]
): XmlAttr[] {
  return attrs.filter(([name]) => !localNames.includes(localPart(name)));
}

export function attrValue(
  attrs: readonly XmlAttr[],
  localName: string
): string | null {
  return attrs.find(([name]) => localPart(name) === localName)?.[1] ?? null;
}

/**
 * An opening tag whose attributes are the ones the original wrote.
 *
 * The text is spliced in exactly as it arrived rather than parsed and written again, so the
 * spelling, the order, and the escaping a producer chose all survive being edited.
 */
export function openTagXml(name: string, rawAttrs: string | null): string {
  return rawAttrs ? `<${name} ${rawAttrs}>` : `<${name}>`;
}

/** The same for an element that closes on its own */
export function emptyTagXml(name: string, rawAttrs: string | null): string {
  return rawAttrs ? `<${name} ${rawAttrs}/>` : `<${name}/>`;
}
