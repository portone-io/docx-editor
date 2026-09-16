/**
 * Spelling the namespaces of an arrived part under the prefixes this package writes.
 * `spec/notes/conformance.md` holds the decision and what it rests on.
 *
 * Names alone are rewritten: element names, attribute names, the namespace declarations
 * themselves, and the prefix tokens of the markup-compatibility attributes that carry a prefix or
 * QName list (§17.2.1-2, §17.2.4-5). No WordprocessingML attribute value holds a QName, so no
 * other value is looked at, and text, CDATA, comments and processing instructions are copied byte
 * for byte - a `w:p` a document wrote inside a `w:t` is its text, not its markup.
 *
 * Nothing here turns a part down. One whose rewrite would not be unambiguous is handed back as it
 * arrived, and what refuses it then is what refused it before this rewrite existed: the reader
 * that needs the prefix (`docx/importDocx`) or the write about to spell it (`ooxml/partSplice`).
 */

import type { XmlAttr } from "./element";
import {
  isKnownPrefix,
  type KnownPrefix,
  NAMESPACES,
  qualify,
  XMLNS,
  xmlnsName,
} from "./names";
import {
  type AttrSpan,
  attrSpans,
  parseAttrs,
  readTag,
  rootTagAt,
  type Tag,
} from "./tagScan";

/** The prefix this package writes each namespace it knows under, keyed by the namespace */
const CANONICAL_PREFIX: ReadonlyMap<string, KnownPrefix> = new Map(
  Object.keys(NAMESPACES)
    .filter(isKnownPrefix)
    .map((prefix) => [NAMESPACES[prefix], prefix])
);

/** The markup-compatibility attributes whose value is a list of prefixes or of QNames */
const MCE_PREFIX_LISTS: ReadonlySet<string> = new Set(
  [
    "Ignorable",
    "MustUnderstand",
    "ProcessContent",
    "PreserveElements",
    "PreserveAttributes",
  ].map((local) => qualify("mc", local))
);

/** The two elements whose unprefixed `Requires` names prefixes rather than a value of its own */
const MCE_ALTERNATE_CONTENT: ReadonlySet<string> = new Set(
  ["Choice", "Fallback"].map((local) => qualify("mc", local))
);

const DECL = `${XMLNS}:`;

/** What a part spells each namespace as, against what this package spells it as */
interface PrefixMap {
  /** The prefix the part wrote, to the prefix this package writes for the same namespace */
  mapped: ReadonlyMap<string, KnownPrefix>;
  /** The prefix unprefixed elements take, for a part whose default namespace is one we know */
  defaultTo: KnownPrefix | null;
  /** The prefixes the rewrite spells out */
  targets: ReadonlySet<KnownPrefix>;
}

/**
 * What the root's declarations ask to be rewritten, and null for a root the rewrite cannot spell
 * unambiguously: a prefix it would write stands there for another namespace already.
 *
 * A namespace this package has no prefix for is left under whatever prefix the part gave it: it is
 * neither read nor written, so its spelling stays the document's own business.
 */
function prefixMapOf(rootAttrs: readonly XmlAttr[]): PrefixMap | null {
  const bound = new Map<string, string>();
  const mapped = new Map<string, KnownPrefix>();
  let defaultTo: KnownPrefix | null = null;
  for (const [name, namespace] of rootAttrs) {
    if (name === XMLNS) {
      defaultTo = CANONICAL_PREFIX.get(namespace) ?? null;
      continue;
    }
    if (!name.startsWith(DECL)) continue;
    const prefix = name.slice(DECL.length);
    bound.set(prefix, namespace);
    const canonical = CANONICAL_PREFIX.get(namespace);
    if (canonical !== undefined && canonical !== prefix)
      mapped.set(prefix, canonical);
  }
  const targets = new Set(mapped.values());
  if (defaultTo !== null) targets.add(defaultTo);
  for (const target of targets) {
    const already = bound.get(target);
    if (already !== undefined && already !== NAMESPACES[target]) return null;
  }
  return { mapped, defaultTo, targets };
}

/** What a name with no prefix stands for, which is a different thing in each of the three places */
type Unprefixed = (name: string, map: PrefixMap) => string;

/** An element name, which the part's default declaration is about */
const underDefault: Unprefixed = (name, map) =>
  map.defaultTo === null ? name : qualify(map.defaultTo, name);

/** An attribute name, which is in no namespace whatever the part declares */
const inNoNamespace: Unprefixed = (name) => name;

/** A compatibility token, which names a prefix rather than carrying a local name of its own */
const asPrefix: Unprefixed = (token, map) => map.mapped.get(token) ?? token;

/** A written name with its prefix swapped for the one this package spells that namespace under */
function mappedName(
  name: string,
  map: PrefixMap,
  unprefixed: Unprefixed
): string {
  const colon = name.indexOf(":");
  if (colon === -1) return unprefixed(name, map);
  const canonical = map.mapped.get(name.slice(0, colon));
  return canonical === undefined ? name : canonical + name.slice(colon);
}

/**
 * The name one attribute goes out under.
 *
 * A declaration is rewritten to the name this package would have declared the same namespace
 * under. `xml` is bound by the XML specification rather than by a declaration, so `xml:space`
 * survives whatever the part declares.
 */
function rewriteAttrName(name: string, map: PrefixMap): string {
  if (name === XMLNS) {
    return map.defaultTo === null ? name : xmlnsName(map.defaultTo);
  }
  if (name.startsWith(DECL)) {
    const canonical = map.mapped.get(name.slice(DECL.length));
    return canonical === undefined ? name : xmlnsName(canonical);
  }
  if (name.startsWith("xml:")) return name;
  return mappedName(name, map, inNoNamespace);
}

function isDeclaration(name: string): boolean {
  return name === XMLNS || name.startsWith(DECL);
}

/** Whether this attribute of this element names prefixes rather than carrying a value of its own */
function namesPrefixes(elementName: string, attrName: string): boolean {
  return (
    MCE_PREFIX_LISTS.has(attrName) ||
    (attrName === "Requires" && MCE_ALTERNATE_CONTENT.has(elementName))
  );
}

/** A prefix or QName list with every token this package spells differently rewritten */
function rewriteTokens(value: string, map: PrefixMap): string {
  return value.replace(/\S+/g, (token) => mappedName(token, map, asPrefix));
}

/** The namespace the rewrite leaves a prefix standing for, and undefined for one it never spells */
function spelledNamespace(prefix: string, map: PrefixMap): string | undefined {
  const canonical = map.mapped.get(prefix);
  if (canonical !== undefined) return NAMESPACES[canonical];
  if (!isKnownPrefix(prefix) || !map.targets.has(prefix)) return undefined;
  return NAMESPACES[prefix];
}

/**
 * Whether a tag binds a prefix this rewrite spells to a namespace other than the one it spells it
 * for, which is what makes the rewrite ambiguous rather than a respelling.
 *
 * Such a part says two things by one prefix, and a name rewritten inside that subtree would mean
 * whichever of them was in scope. A declaration naming the namespace the prefix already stands for
 * says nothing new and is rewritten like the root's.
 * (`ooxml/fragment` guards a pasted fragment against rebinding `w` or `r` at all; this asks only
 * whether the prefixes it is about to spell are unambiguous throughout the part.)
 */
function rebindsSpelledPrefix(
  raw: string,
  spans: readonly AttrSpan[],
  map: PrefixMap
): boolean {
  return spans.some((span) => {
    const name = raw.slice(span.nameStart, span.nameEnd);
    if (!isDeclaration(name)) return false;
    const namespace = raw.slice(span.valueStart, span.valueEnd);
    if (name === XMLNS) {
      return map.defaultTo !== null && namespace !== NAMESPACES[map.defaultTo];
    }
    const spelled = spelledNamespace(name.slice(DECL.length), map);
    return spelled !== undefined && namespace !== spelled;
  });
}

/**
 * The attribute text of one tag, rewritten in place, and null for a tag the rewrite cannot be
 * carried through: one leaving a prefix it spells ambiguous, and one whose attributes cannot be
 * read at all.
 *
 * Every byte between the spans is copied, so the quoting, the spacing and the spelling of a
 * reference stay as the part wrote them. Two declarations of one namespace collapse into the
 * single one this package writes, since a tag cannot carry the same name twice.
 */
function rewriteAttrs(
  raw: string,
  elementName: string,
  map: PrefixMap
): string | null {
  const spans = attrSpans(raw);
  if (spans === null) return null;
  if (raw.includes(XMLNS) && rebindsSpelledPrefix(raw, spans, map)) return null;
  const declared = new Set(
    spans.map((span) => raw.slice(span.nameStart, span.nameEnd))
  );
  const written = new Set<string>();
  let out = "";
  let at = 0;
  for (const span of spans) {
    const name = raw.slice(span.nameStart, span.nameEnd);
    const rewritten = rewriteAttrName(name, map);
    if (
      rewritten !== name &&
      isDeclaration(rewritten) &&
      (declared.has(rewritten) || written.has(rewritten))
    ) {
      at = span.valueEnd + 1;
      continue;
    }
    written.add(rewritten);
    const value = raw.slice(span.valueStart, span.valueEnd);
    out +=
      raw.slice(at, span.nameStart) +
      rewritten +
      raw.slice(span.nameEnd, span.valueStart) +
      (namesPrefixes(elementName, rewritten)
        ? rewriteTokens(value, map)
        : value) +
      span.quote;
    at = span.valueEnd + 1;
  }
  return out + raw.slice(at);
}

/** One element tag with its name and its attributes rewritten, its shape left as it stood */
function rewriteTag(xml: string, tag: Tag, map: PrefixMap): string | null {
  const closeLength = tag.kind === "empty" ? 2 : 1;
  const name = mappedName(tag.name, map, underDefault);
  const raw = xml.slice(tag.nameEnd, tag.end - closeLength);
  const attrs = tag.kind === "close" ? raw : rewriteAttrs(raw, name, map);
  if (attrs === null) return null;
  return (
    (tag.kind === "close" ? "</" : "<") +
    name +
    attrs +
    xml.slice(tag.end - closeLength, tag.end)
  );
}

/**
 * The part text with every namespace this package knows spelled the way it writes it.
 *
 * A part that comes back unrewritten comes back as the same string by identity, so a package no
 * producer wrote foreign prefixes into is never re-encoded and still exports byte for byte.
 */
export function withEditorPrefixes(xml: string): string {
  const rootAt = rootTagAt(xml);
  if (rootAt === -1) return xml;
  const root = readTag(xml, rootAt);
  if (root === null || root.kind === "close") return xml;
  const rootAttrs = parseAttrs(
    xml.slice(root.nameEnd, root.end - (root.kind === "empty" ? 2 : 1))
  );
  if (rootAttrs === null) return xml;
  const map = prefixMapOf(rootAttrs);
  if (map === null) return xml;
  if (map.mapped.size === 0 && map.defaultTo === null) return xml;

  const out: string[] = [];
  let at = 0;
  for (;;) {
    const lt = xml.indexOf("<", at);
    if (lt === -1) break;
    const tag = readTag(xml, lt);
    if (tag === null) return xml;
    const rewritten =
      tag.kind === "other" ? xml.slice(lt, tag.end) : rewriteTag(xml, tag, map);
    if (rewritten === null) return xml;
    out.push(xml.slice(at, lt));
    out.push(rewritten);
    at = tag.end;
  }
  out.push(xml.slice(at));
  return out.join("");
}
