/**
 * The shapes this editor writes the three comment parts in, and the reading of those shapes.
 *
 * Each entry is written here and judged here, so the two halves cannot drift apart: a body written
 * by hand can be told from one this editor put out only where the writer and the reader agree on
 * what it puts out. `./reading` keeps a lenient flattener for display, which takes whatever a
 * producer wrote and shows what it can; this one judges an entry that came back and says no to
 * everything else.
 */

import {
  attrsText,
  elementXml,
  type XmlAttr,
  xmlnsAttr,
} from "../../ooxml/element";
import { qualify, wName } from "../../ooxml/names";
import {
  attributeByLocalName,
  elementChildren,
  escapeXml,
  parseXml,
  W_NS,
} from "../../ooxml/xml";
import { COMMENT_AUTHOR_PROVIDER, W14_NS, W15_NS } from "./constants";
import type { CommentReferenceData, CommentReplyData } from "./model";

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

const XMLNS_NS = "http://www.w3.org/2000/xmlns/";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

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

const PARAGRAPH_ATTRIBUTES: ReadonlySet<string> = new Set([
  nameKey(W14_NS, "paraId"),
]);

const TEXT_ATTRIBUTES: ReadonlySet<string> = new Set([
  nameKey(XML_NS, "space"),
]);

const PERSON_ATTRIBUTES: ReadonlySet<string> = new Set([
  nameKey(W15_NS, "author"),
]);

const PRESENCE_ATTRIBUTES: ReadonlySet<string> = new Set([
  nameKey(W15_NS, "providerId"),
  nameKey(W15_NS, "userId"),
]);

const NO_ATTRIBUTES: ReadonlySet<string> = new Set();

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
 * Whether the element carries no attribute outside this set.
 *
 * A namespace declaration is not an attribute of the element in this sense: it says where the
 * names come from rather than anything about the entry, and where a producer puts one is its own
 * business.
 */
export function attributesWithin(
  el: Element,
  allowed: ReadonlySet<string>
): boolean {
  return Array.from(el.attributes).every(
    (attr) =>
      attr.namespaceURI === XMLNS_NS ||
      attr.name === "xmlns" ||
      allowed.has(nameKey(attr.namespaceURI, attr.localName))
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
function holdsElementsOnly(el: Element): boolean {
  return Array.from(el.childNodes).every(
    (node) => node.nodeType === ELEMENT_NODE
  );
}

/** Whether the element holds text and nothing else, which is what a `w:t` holds */
function holdsTextOnly(el: Element): boolean {
  return Array.from(el.childNodes).every((node) => node.nodeType === TEXT_NODE);
}

/** The body of a comment, as a paragraph of one run holding the text and the breaks in it */
export function renderCommentBody(text: string, paraId: string | null): string {
  const lines = text.split("\n");
  const pieces: string[] = [];
  lines.forEach((line, index) => {
    if (index > 0) pieces.push(elementXml(wName("br"), []));
    if (line.length > 0 || lines.length === 1) {
      pieces.push(
        elementXml(wName("t"), [["xml:space", "preserve"]], [escapeXml(line)])
      );
    }
  });
  const key: readonly XmlAttr[] =
    paraId === null
      ? []
      : [xmlnsAttr("w14"), [qualify("w14", "paraId"), paraId]];
  return elementXml(wName("p"), key, [elementXml(wName("r"), [], pieces)]);
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

function readRunText(run: Element): string | null {
  if (!holdsElementsOnly(run)) return null;
  const pieces: string[] = [];
  for (const child of Array.from(run.children)) {
    if (isNamed(child, W_NS, "br")) {
      if (
        !attributesWithin(child, NO_ATTRIBUTES) ||
        child.childNodes.length > 0
      ) {
        return null;
      }
      pieces.push("\n");
      continue;
    }
    if (!isNamed(child, W_NS, "t")) return null;
    // The writer keeps the space of every line, so it says so on every `w:t` it writes
    if (
      !attributesWithin(child, TEXT_ATTRIBUTES) ||
      child.getAttributeNS(XML_NS, "space") !== "preserve" ||
      !holdsTextOnly(child)
    ) {
      return null;
    }
    pieces.push(child.textContent ?? "");
  }
  return pieces.join("");
}

/**
 * The text of a comment written in that shape, and null for an entry holding anything else: a
 * field, a second run, markup a producer wrapped it in, an attribute this editor does not write.
 *
 * An entry that arrived and was not edited is compared as it stands rather than read here, so
 * saying no to a shape this editor would not have written turns down only a rewrite.
 */
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

export function readStrictCommentBody(comment: Element): string | null {
  if (!holdsElementsOnly(comment)) return null;
  const paragraphs = Array.from(comment.children);
  if (paragraphs.length !== 1) return null;
  const [paragraph] = paragraphs;
  if (!isNamed(paragraph, W_NS, "p")) return null;
  if (!attributesWithin(paragraph, PARAGRAPH_ATTRIBUTES)) return null;

  if (!holdsElementsOnly(paragraph)) return null;
  const runs = Array.from(paragraph.children);
  if (runs.length !== 1) return null;
  const [run] = runs;
  if (!isNamed(run, W_NS, "r")) return null;
  if (!attributesWithin(run, NO_ATTRIBUTES)) return null;
  return readRunText(run);
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
