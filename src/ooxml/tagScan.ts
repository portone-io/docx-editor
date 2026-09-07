/**
 * Reading one tag out of a part or a fragment without parsing the whole of it.
 *
 * Everything the writers hold is text, and an edit slices one element or one attribute out of it,
 * so the spots those slices run between have to be found in the text itself. A `>` may stand
 * inside an attribute value and a `<` inside a comment, which is why a tag is read construct by
 * construct rather than searched for.
 */

import type { XmlAttr } from "./element";

export type TagKind = "open" | "close" | "empty" | "other";

export interface Tag {
  kind: TagKind;
  /** The name exactly as written, prefix included. Empty for a construct that has none */
  name: string;
  /** The spot where the tag name ends. This is the start of the attribute string */
  nameEnd: number;
  /** The spot right after the tag ends */
  end: number;
}

function skipPast(source: string, from: number, marker: string): number {
  const at = source.indexOf(marker, from);
  return at === -1 ? -1 : at + marker.length;
}

/** A construct that is not an element tag: a declaration, a comment, a CDATA section, a doctype */
function otherTag(end: number): Tag | null {
  return end === -1 ? null : { kind: "other", name: "", nameEnd: end, end };
}

/** Reads a single tag starting at a `<`. null if it cannot be made out */
export function readTag(source: string, lt: number): Tag | null {
  if (source.startsWith("<?", lt)) return otherTag(skipPast(source, lt, "?>"));
  if (source.startsWith("<!--", lt))
    return otherTag(skipPast(source, lt, "-->"));
  if (source.startsWith("<![CDATA[", lt))
    return otherTag(skipPast(source, lt, "]]>"));
  if (source.startsWith("<!", lt)) return otherTag(skipPast(source, lt, ">"));

  const closing = source.startsWith("</", lt);
  let i = lt + (closing ? 2 : 1);
  const nameStart = i;
  while (i < source.length && !" \t\r\n/>".includes(source[i])) i += 1;
  const name = source.slice(nameStart, i);
  if (!name) return null;
  const nameEnd = i;

  let quote: string | null = null;
  while (i < source.length) {
    const ch = source[i];
    if (quote !== null) {
      if (ch === quote) quote = null;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === ">") {
      const selfClosing = source[i - 1] === "/";
      const kind: TagKind = closing ? "close" : selfClosing ? "empty" : "open";
      return { kind, name, nameEnd, end: i + 1 };
    }
    i += 1;
  }
  return null;
}

/**
 * Where the root element's opening tag starts, past the prolog.
 *
 * The prolog is walked construct by construct so that a `<` inside a comment or a processing
 * instruction is not taken for the root. -1 when the source holds no element at all.
 */
export function rootTagAt(xml: string): number {
  let at = 0;
  for (;;) {
    const lt = xml.indexOf("<", at);
    if (lt === -1) return -1;
    const tag = readTag(xml, lt);
    if (!tag) return -1;
    if (tag.kind !== "other") return lt;
    at = tag.end;
  }
}

/** The references a part may carry with no DTD to declare any of its own */
const ENTITIES: Readonly<Record<string, string>> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

const HIGHEST_CODE_POINT = 0x10ffff;

/** What one reference stands for, and null for one no parser would resolve */
function referenceText(name: string): string | null {
  const named = ENTITIES[name];
  if (named !== undefined) return named;
  const numeric = /^#(?:[xX]([0-9a-fA-F]+)|(\d+))$/.exec(name);
  if (!numeric) return null;
  const [, hex, decimal] = numeric;
  const code = Number.parseInt(hex ?? decimal, hex === undefined ? 10 : 16);
  // A surrogate half and a code point past the last plane are both characters XML cannot carry
  if (code > HIGHEST_CODE_POINT || (code >= 0xd800 && code <= 0xdfff)) {
    return null;
  }
  return String.fromCodePoint(code);
}

/** The text an attribute value stands for. null when it carries a reference nothing declares */
function decodeReferences(value: string): string | null {
  if (!value.includes("&")) return value;
  let decoded = "";
  let at = 0;
  for (;;) {
    const amp = value.indexOf("&", at);
    if (amp === -1) return decoded + value.slice(at);
    const semicolon = value.indexOf(";", amp);
    if (semicolon === -1) return null;
    const text = referenceText(value.slice(amp + 1, semicolon));
    if (text === null) return null;
    decoded += value.slice(at, amp) + text;
    at = semicolon + 1;
  }
}

/**
 * One attribute and the whitespace in front of it. A quoted value may hold neither the quote that
 * opened it nor a `<`, both of which end the value for any parser.
 */
const ATTRIBUTE = /(\s*)([^\s=/><"']+)\s*=\s*(?:"([^"<]*)"|'([^'<]*)')/y;

/**
 * The attribute text of an opening tag as pairs, references decoded, each name spelled as it was
 * written. null for a shape a parser would refuse: an unquoted value, a missing name, one name
 * given twice, or a reference nothing declares.
 */
export function parseAttrs(raw: string): XmlAttr[] | null {
  const attrs: XmlAttr[] = [];
  const names = new Set<string>();
  let at = 0;
  for (;;) {
    ATTRIBUTE.lastIndex = at;
    const matched = ATTRIBUTE.exec(raw);
    if (!matched) return raw.slice(at).trim() === "" ? attrs : null;
    const [, separator, name, doubled, singled] = matched;
    // Two attributes have to stand apart, and the first one starts the text
    if (at > 0 && separator === "") return null;
    if (names.has(name)) return null;
    names.add(name);
    const value = decodeReferences(doubled ?? singled);
    if (value === null) return null;
    attrs.push([name, value]);
    at = ATTRIBUTE.lastIndex;
  }
}
