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

const XMLNS_NS = "http://www.w3.org/2000/xmlns/";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

const nameKey = (namespace: string | null, localName: string): string =>
  `${namespace ?? ""} ${localName}`;

/** The attributes this editor writes on a `w:comment` */
export const COMMENT_ATTRIBUTES: ReadonlySet<string> = new Set([
  nameKey(W_NS, "id"),
  nameKey(W_NS, "author"),
  nameKey(W_NS, "date"),
  nameKey(W_NS, "initials"),
  nameKey(W14_NS, "paraId"),
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

function readRunText(run: Element): string | null {
  const pieces: string[] = [];
  for (const child of Array.from(run.children)) {
    if (isNamed(child, W_NS, "br")) {
      if (
        !attributesWithin(child, NO_ATTRIBUTES) ||
        child.children.length > 0
      ) {
        return null;
      }
      pieces.push("\n");
      continue;
    }
    if (!isNamed(child, W_NS, "t")) return null;
    if (
      !attributesWithin(child, TEXT_ATTRIBUTES) ||
      child.children.length > 0
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
  const paragraphs = Array.from(comment.children);
  if (paragraphs.length !== 1) return null;
  const [paragraph] = paragraphs;
  if (!isNamed(paragraph, W_NS, "p")) return null;
  if (!attributesWithin(paragraph, PARAGRAPH_ATTRIBUTES)) return null;

  const runs = Array.from(paragraph.children);
  if (runs.length !== 1) return null;
  const [run] = runs;
  if (!isNamed(run, W_NS, "r")) return null;
  if (!attributesWithin(run, NO_ATTRIBUTES)) return null;
  return readRunText(run);
}
