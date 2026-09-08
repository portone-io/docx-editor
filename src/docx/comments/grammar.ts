/**
 * The shapes this editor writes the three comment parts in, and the reading of those shapes.
 *
 * Each entry is written here and judged here, so the two halves cannot drift apart: a body written
 * by hand can be told from one this editor put out only where the writer and the reader agree on
 * what it puts out. `./reading` keeps a lenient flattener for display, which takes whatever a
 * producer wrote and shows what it can; this one judges an entry that came back and says no to
 * everything else.
 */

import { attrsText, elementXml, type XmlAttr } from "../../ooxml/element";
import { NAMESPACES, qualify } from "../../ooxml/names";
import {
  attributeByLocalName,
  elementChildren,
  parseXml,
  W_NS,
} from "../../ooxml/xml";
import { COMMENT_AUTHOR_PROVIDER, W14_NS, W15_NS } from "./constants";
import type { CommentReferenceData, CommentReplyData } from "./model";

const ELEMENT_NODE = 1;

const TEXT_NODE = 3;

const XMLNS_NS = "http://www.w3.org/2000/xmlns/";

const nameKey = (namespace: string | null, localName: string): string =>
  `${namespace ?? ""} ${localName}`;

/** The attributes this editor writes on a `w:comment`. The thread key goes on the body's paragraph */
export const COMMENT_ATTRIBUTES: ReadonlySet<string> = new Set([
  nameKey(W_NS, "id"),
  nameKey(W_NS, "author"),
  nameKey(W_NS, "date"),
  nameKey(W_NS, "initials"),
]);

/** The attributes this editor writes on a `w15:commentEx` */
export const COMMENT_EX_ATTRIBUTES: ReadonlySet<string> = new Set([
  nameKey(W15_NS, "paraId"),
  nameKey(W15_NS, "paraIdParent"),
  nameKey(W15_NS, "done"),
]);

const PERSON_ATTRIBUTES: ReadonlySet<string> = new Set([
  nameKey(W15_NS, "author"),
]);

const PRESENCE_ATTRIBUTES: ReadonlySet<string> = new Set([
  nameKey(W15_NS, "providerId"),
  nameKey(W15_NS, "userId"),
]);

/**
 * A thread key, which is four bytes written as hexadecimal (`ST_LongHexNumber`, ECMA-376 Part 1
 * §17.18.51). The writer takes its own from `nextCommentParaId`, and a key that arrived was
 * written by a producer holding to the same type.
 */
const THREAD_KEY = /^[0-9A-Fa-f]{8}$/;

/** Whether the attribute is absent, or a thread key */
function threadKeyOrNone(el: Element, localName: string): boolean {
  const value = attributeByLocalName(el, localName);
  return value === null || THREAD_KEY.test(value);
}

/**
 * The namespaces this package writes a declaration for, each under the prefix it writes it under.
 *
 * The writer declares a namespace on an entry it builds itself and nowhere else: `w` on a comment,
 * `w14` on the paragraph a thread key goes on, `w15` on an identity written into a part whose root
 * binds nothing.
 */
const WRITTEN_NAMESPACES: ReadonlyMap<string, string> = new Map([
  ["w", NAMESPACES.w],
  ["w14", NAMESPACES.w14],
  ["w15", NAMESPACES.w15],
]);

/**
 * Whether the namespace declaration is one this editor's writer puts out, its value included.
 *
 * A declaration decides what every name around it means, so one binding a prefix to a namespace
 * this package does not write is markup the writer did not put there however ordinary the names
 * under it look. The default declaration is never one of them: the writer spells out a prefix on
 * everything it writes.
 */
export function declarationWritten(attr: Attr): boolean {
  if (!attr.name.startsWith("xmlns:")) return false;
  return (
    WRITTEN_NAMESPACES.get(attr.name.slice("xmlns:".length)) === attr.value
  );
}

/**
 * Whether this element and everything inside it declares only namespaces this package writes.
 *
 * A declaration decides what every name under it means, so a block rebinding a prefix says one
 * thing to this reader and another to Word, and is a place to put bytes nothing looks at.
 */
export function declarationsWritten(el: Element): boolean {
  return (
    Array.from(el.attributes).every(
      (attr) => !attr.name.startsWith("xmlns") || declarationWritten(attr)
    ) && Array.from(el.children).every(declarationsWritten)
  );
}

/**
 * Whether the element carries no attribute outside this set, and declares no namespace outside
 * what the writer declares.
 */
export function attributesWithin(
  el: Element,
  allowed: ReadonlySet<string>
): boolean {
  return Array.from(el.attributes).every((attr) =>
    attr.namespaceURI === XMLNS_NS
      ? declarationWritten(attr)
      : allowed.has(nameKey(attr.namespaceURI, attr.localName))
  );
}

function isNamed(el: Element, namespace: string, localName: string): boolean {
  return el.namespaceURI === namespace && el.localName === localName;
}

/**
 * Whether the element holds elements and nothing else.
 *
 * The writer puts no whitespace, comment or stray text between the pieces of a body, so an entry
 * carrying any is one it did not write, and a reading that passed over them would leave a place to
 * put bytes nothing looks at.
 */
export function holdsElementsOnly(el: Element): boolean {
  return Array.from(el.childNodes).every(
    (node) => node.nodeType === ELEMENT_NODE
  );
}

/** XML layout whitespace, which is the one text carrying no payload wherever it stands */
export function isLayoutText(node: Node): boolean {
  return (
    node.nodeType === TEXT_NODE && /^[ \t\r\n]*$/.test(node.nodeValue ?? "")
  );
}

/**
 * Whether the element holds elements and the whitespace a producer laid them out over.
 *
 * A block is sliced with the whitespace standing ahead of it (`docx/scan`), so the writer hands a
 * block it left untouched back with the layout it arrived in, and an entry from a part written
 * over several lines carries that layout between its blocks. Whitespace is the only text this
 * passes over: a comment, a processing instruction or a stretch of real text is a place to put
 * bytes nothing looks at, and none of them is anything the writer puts out.
 */
export function holdsElementsAndLayout(el: Element): boolean {
  return Array.from(el.childNodes).every(
    (node) => node.nodeType === ELEMENT_NODE || isLayoutText(node)
  );
}

/**
 * An opening tag named `p`, or a stretch of text that only looks like one.
 *
 * A prefix is an XML name, which is wider than the letters an English one uses, so it is spelled
 * here as everything a tag cannot hold rather than as the characters one usually does. The DOM
 * counts the same tags, and the two have to agree on which they are.
 */
const PARAGRAPH_OR_SKIPPED =
  /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<([^\s<>/:="']+:)?p(?=[\s/>])[^>]*>/g;

/**
 * The paragraph a thread key belongs on, which is the last one of the body.
 *
 * A body may hold a paragraph of another vocabulary, a picture's DrawingML `a:p` among them, and a
 * key put on that says nothing about the thread. The namespace decides, since a prefix means only
 * what the element it sits on binds it to, and a document may bind the same prefix twice over.
 */
export function lastBodyParagraph(comment: Element): Element | null {
  const paragraphs = comment.getElementsByTagNameNS(W_NS, "p");
  return paragraphs.length === 0 ? null : paragraphs[paragraphs.length - 1];
}

/** Every element of the body whose name is `p`, whatever it means, in the order the text has them */
function namedParagraphs(comment: Element): readonly Element[] {
  return Array.from(comment.getElementsByTagName("*")).filter(
    (element) => element.localName === "p"
  );
}

/**
 * The entries a part arrived holding, under the id each carries.
 *
 * The writer asks the document which paragraph a key goes on rather than asking the text, because
 * only the document knows what a prefix means where it is written.
 */
export function arrivedEntries(
  partXml: string | null
): ReadonlyMap<string, Element> {
  const entries = new Map<string, Element>();
  if (partXml === null) return entries;
  let root: Element;
  try {
    root = parseXml(partXml).documentElement;
  } catch {
    return entries;
  }
  for (const entry of elementChildren(root)) {
    if (entry.namespaceURI !== W_NS || entry.localName !== "comment") continue;
    const id = attributeByLocalName(entry, "id");
    if (id !== null && !entries.has(id)) entries.set(id, entry);
  }
  return entries;
}

/**
 * The entry with the thread key on its body's last paragraph, and unchanged where it has one.
 *
 * Settling a thread or replying to it hangs the state off that key, and an entry that arrived
 * without one has to gain it. Writing the entry afresh instead would put back only what this
 * editor models, so a body holding more than plain text would lose it to a change nobody asked
 * for and nobody made.
 *
 * `arrived` is that entry as the document has it, which is what says which paragraph is a
 * WordprocessingML one and whether it already carries a key under some prefix. The text is only
 * asked where that paragraph is: its opening tag is the one at the same place among the tags named
 * `p`, counted the same way in both.
 *
 * The `w14` prefix is declared on the part rather than here: a part holding any thread state
 * declares it on its root along with the compatibility markup that goes with it (`./writing`).
 */
export function withThreadKey(
  commentXml: string,
  paraId: string,
  arrived: Element | null
): string {
  if (arrived === null) return commentXml;
  const target = lastBodyParagraph(arrived);
  if (target === null) return commentXml;
  // Under the namespace it means, or under the name the key would be written as: a second
  // attribute of one name is not XML this package can read back
  if (
    target.getAttributeNS(W14_NS, "paraId") !== null ||
    target.hasAttribute("w14:paraId")
  ) {
    return commentXml;
  }

  const at = namedParagraphs(arrived).indexOf(target);
  const openings = Array.from(commentXml.matchAll(PARAGRAPH_OR_SKIPPED)).filter(
    (match) => !match[0].startsWith("<!")
  );
  const opening = openings[at];
  if (at === -1 || opening === undefined || opening.index === undefined) {
    return commentXml;
  }
  const selfClosing = opening[0].endsWith("/>");
  const keyed =
    opening[0].slice(0, selfClosing ? -2 : -1) +
    ` ${attrsText([[qualify("w14", "paraId"), paraId]])}` +
    (selfClosing ? "/>" : ">");
  return (
    commentXml.slice(0, opening.index) +
    keyed +
    commentXml.slice(opening.index + opening[0].length)
  );
}

/** Whether this editor's writer could have put out this thread state */
export function wellFormedCommentExtension(entry: Element): boolean {
  return (
    attributesWithin(entry, COMMENT_EX_ATTRIBUTES) &&
    entry.childNodes.length === 0 &&
    threadKeyOrNone(entry, "paraId") &&
    threadKeyOrNone(entry, "paraIdParent") &&
    // The writer says a thread is settled or open, and nothing else `ST_OnOff` would take
    ["0", "1", null].includes(attributeByLocalName(entry, "done"))
  );
}

/** Whether this editor's writer could have put out this recorded identity */
export function wellFormedPerson(entry: Element): boolean {
  if (!attributesWithin(entry, PERSON_ATTRIBUTES)) return false;
  if (!holdsElementsOnly(entry)) return false;
  const children = Array.from(entry.children);
  if (children.length !== 1) return false;
  const [presence] = children;
  return (
    isNamed(presence, W15_NS, "presenceInfo") &&
    attributesWithin(presence, PRESENCE_ATTRIBUTES) &&
    attributeByLocalName(presence, "providerId") === COMMENT_AUTHOR_PROVIDER &&
    attributeByLocalName(presence, "userId") !== null &&
    presence.childNodes.length === 0
  );
}

/** The identity a recorded person stands for, and null for an entry recording none */
export function recordedIdentity(person: Element): string | null {
  const [presence] = Array.from(person.children);
  return presence === undefined
    ? null
    : attributeByLocalName(presence, "userId");
}

/** The thread state of one comment, as this editor writes it into the extended part */
export function renderCommentExtension(
  comment: CommentReferenceData | CommentReplyData
): string {
  if (comment.extensionXml !== null) return comment.extensionXml;
  const parent: readonly XmlAttr[] =
    "parentParaId" in comment
      ? [[qualify("w15", "paraIdParent"), comment.parentParaId]]
      : [];
  const done: readonly XmlAttr[] =
    "resolved" in comment
      ? [[qualify("w15", "done"), comment.resolved ? "1" : "0"]]
      : [];
  return elementXml(qualify("w15", "commentEx"), [
    [qualify("w15", "paraId"), comment.paraId],
    ...parent,
    ...done,
  ]);
}

/**
 * The identity this editor records for an author, as it writes it into the people part.
 *
 * The prefix and the declaration come from the caller, which is writing into a part whose root
 * already binds them or is being written from nothing.
 */
export function renderPerson(
  author: string,
  userId: string,
  prefix: string,
  declaration: XmlAttr | null
): string {
  const declared: readonly XmlAttr[] =
    declaration === null ? [] : [declaration];
  return elementXml(
    `${prefix}person`,
    [...declared, [`${prefix}author`, author]],
    [
      elementXml(`${prefix}presenceInfo`, [
        [`${prefix}providerId`, COMMENT_AUTHOR_PROVIDER],
        [`${prefix}userId`, userId],
      ]),
    ]
  );
}
