/**
 * Writing one element out as text, and merging the attributes of one.
 *
 * Everything here works on strings, because that is what the writers hold: a fragment the file
 * arrived with is carried around as its original text, and an edit swaps out one element or one
 * attribute of it. Rebuilding a parsed element instead would rewrite markup nobody touched.
 *
 * An attribute is read and written as the WordprocessingML attribute of a local name: the one
 * spelled under the `w` prefix, which `RESERVED_PREFIXES` pins to that namespace in every fragment,
 * or failing that an unprefixed one, which is what `wAttr` accepted off a parsed fragment too. An
 * attribute another prefix qualifies belongs to some other vocabulary, so it is neither read as the
 * formatting value nor written over when it happens to share the local name (`x:val` beside
 * `w:val`). The two readers below are named for which of those they answer.
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
 * The local name this written name stands for as a WordprocessingML attribute: `w:val` and `val`
 * both answer `val`. null for a name another prefix qualifies, which is not one of ours to read
 * or to write.
 */
export function wLocalName(name: string): string | null {
  if (!name.includes(":")) return name;
  const local = localPart(name);
  return name === wName(local) ? local : null;
}

/** Where the WordprocessingML attribute of this local name sits, the `w:` spelling ahead of an unprefixed one. -1 for none */
function wAttrIndex(attrs: readonly XmlAttr[], localName: string): number {
  const qualified = attrs.findIndex(([name]) => name === wName(localName));
  if (qualified !== -1) return qualified;
  return attrs.findIndex(([name]) => name === localName);
}

/**
 * The value of the WordprocessingML attribute of this local name, and null for an element that
 * carries none. An attribute of the same local name under another prefix is not it.
 */
export function wAttrValue(
  attrs: readonly XmlAttr[],
  localName: string
): string | null {
  const at = wAttrIndex(attrs, localName);
  return at === -1 ? null : attrs[at][1];
}

/**
 * The attributes with the WordprocessingML one of this local name written as `value`, or removed
 * when it is null.
 *
 * One already there keeps the slot it sat in and the name it was written under, so an edited
 * element reads as the original did. One that is not there yet goes on the end under the `w`
 * prefix, and one of the same local name under another prefix stands untouched beside it.
 */
export function withAttr(
  attrs: readonly XmlAttr[],
  name: string,
  value: string | null
): XmlAttr[] {
  if (value === null) return withoutAttrs(attrs, [name]);
  const at = wAttrIndex(attrs, name);
  if (at === -1) return [...attrs, [wName(name), value]];
  return attrs.map(
    (attr, index): XmlAttr => (index === at ? [attr[0], value] : attr)
  );
}

/** The attributes without the WordprocessingML ones of these local names, in either spelling */
export function withoutAttrs(
  attrs: readonly XmlAttr[],
  localNames: readonly string[]
): XmlAttr[] {
  return attrs.filter(([name]) => {
    const local = wLocalName(name);
    return local === null || !localNames.includes(local);
  });
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
