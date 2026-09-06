/**
 * The one shape this editor writes a comment in, and the reading of that shape.
 *
 * The writer settles on a single grammar, so a body written here can be told from one written by
 * hand: a paragraph holding a single run of text and line breaks, and nothing besides. `./reading`
 * keeps a lenient flattener for display, which takes whatever a producer wrote and shows what it
 * can; this one judges an entry that came back and has to say no to everything else.
 */

import { escapeXml, W_NS } from "../../ooxml/xml";
import { W14_NS, W15_NS } from "./constants";

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

const NO_ATTRIBUTES: ReadonlySet<string> = new Set();

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
    if (index > 0) pieces.push("<w:br/>");
    if (line.length > 0 || lines.length === 1) {
      pieces.push(`<w:t xml:space="preserve">${escapeXml(line)}</w:t>`);
    }
  });
  const attrs =
    paraId === null
      ? ""
      : ` xmlns:w14="${W14_NS}" w14:paraId="${escapeXml(paraId)}"`;
  return `<w:p${attrs}><w:r>${pieces.join("")}</w:r></w:p>`;
}

/** The opening tag of the last paragraph of a fragment, which is where the thread key goes */
const LAST_PARAGRAPH = /<([\w.-]+:)?p(?=[\s/>])[^>]*>/g;

/**
 * The entry with the thread key on its body's last paragraph, and unchanged where it has one.
 *
 * Settling a thread or replying to it hangs the state off that key, and an entry that arrived
 * without one has to gain it. Writing the entry afresh instead would put back only what this
 * editor models, so a body holding more than plain text would lose it to a change nobody asked
 * for and nobody made.
 *
 * The prefix is declared on the part rather than here: a part holding any thread state declares it
 * on its root along with the compatibility markup that goes with it (`./writing`).
 */
export function withThreadKey(commentXml: string, paraId: string): string {
  const openings = Array.from(commentXml.matchAll(LAST_PARAGRAPH));
  const last = openings[openings.length - 1];
  if (last === undefined || last.index === undefined) return commentXml;
  if (/\sw14:paraId\s*=/.test(last[0])) return commentXml;
  const selfClosing = last[0].endsWith("/>");
  const opening =
    last[0].slice(0, selfClosing ? -2 : -1) +
    ` w14:paraId="${escapeXml(paraId)}"` +
    (selfClosing ? "/>" : ">");
  return (
    commentXml.slice(0, last.index) +
    opening +
    commentXml.slice(last.index + last[0].length)
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
