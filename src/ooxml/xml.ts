import { DocxImportError } from "./errors";

/** The wordprocessing namespace that every element we read lives in */
export const W_NS =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * The relationship namespace. It is both where the `r:embed` attributes live and the base
 * every relationship type name is built on
 */
export const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The name with its namespace prefix stripped off (`w:ascii` -> `ascii`) */
export function localPart(name: string): string {
  const colon = name.indexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

/**
 * The control characters XML 1.0 cannot carry.
 * If even one of them makes it inside, Word will fail to open the file.
 */
// biome-ignore lint/suspicious/noControlCharactersInRegex: stripping this very list is what this rule is for
const FORBIDDEN = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/;
const ALL_FORBIDDEN = new RegExp(FORBIDDEN, "g");

/**
 * Makes a value safe to drop verbatim into either text content or an attribute value.
 *
 * Attribute values are wrapped in double quotes everywhere we write them, so single quotes are left alone.
 * Control characters that cannot be carried are stripped. Incoming text is already filtered by
 * `editor/plainText`, but this is the last net that keeps an unopenable file from going out
 * no matter which path the text came in through.
 */
export function escapeXml(value: string): string {
  const escaped = value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
  return FORBIDDEN.test(escaped) ? escaped.replace(ALL_FORBIDDEN, "") : escaped;
}

export function decodeUtf8(bytes: Uint8Array): {
  text: string;
  hadBom: boolean;
} {
  const hadBom =
    bytes.length >= 3 &&
    bytes[0] === 0xef &&
    bytes[1] === 0xbb &&
    bytes[2] === 0xbf;
  return { text: new TextDecoder("utf-8").decode(bytes), hadBom };
}

export function encodeUtf8(text: string, withBom: boolean): Uint8Array {
  return new TextEncoder().encode(withBom ? "\u{FEFF}" + text : text);
}

/**
 * Whether the source declares a DTD. Only the prolog is scanned, walking construct by construct so
 * a comment or processing instruction carrying a `<` (e.g. `<!-- < -->`) cannot end the scan early
 * and hide a following doctype, and a `<!DOCTYPE` in text or CDATA is not mistaken for one.
 */
function declaresDtd(source: string): boolean {
  const DOCTYPE = "<!DOCTYPE";
  let at = 0;
  for (;;) {
    const opens = source.indexOf("<", at);
    if (opens === -1) return false;
    const opening = source.slice(opens, opens + DOCTYPE.length);
    if (opening.toUpperCase() === DOCTYPE) return true;
    // Anything that is neither a comment nor a processing instruction is the root element
    const closes = opening.startsWith("<!--")
      ? source.indexOf("-->", opens + 4)
      : opening.startsWith("<?")
        ? source.indexOf("?>", opens + 2)
        : -1;
    // An unterminated one is markup the parser refuses anyway, so nothing is left to read
    if (closes === -1) return false;
    at = closes;
  }
}

export function parseXml(source: string): Document {
  // ECMA-376 allows no DTD in a package part, and DOMParser expands the entities one
  // declares, so a part carrying one could show text that the part itself does not hold
  if (declaresDtd(source)) {
    throw new DocxImportError("malformed-xml", "the XML declares a DTD");
  }
  const doc = new DOMParser().parseFromString(source, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) {
    throw new DocxImportError("malformed-xml", "could not parse the XML");
  }
  return doc;
}

export function elementChildren(el: Element): Element[] {
  return Array.from(el.children);
}

export function childByLocalName(el: Element, name: string): Element | null {
  return elementChildren(el).find((child) => child.localName === name) ?? null;
}

/** Turns an opening tag's attributes into an `a="b" c="d"` string */
export function attrString(el: Element): string | null {
  if (el.attributes.length === 0) return null;
  return Array.from(el.attributes)
    .map((attr) => `${attr.name}="${escapeXml(attr.value)}"`)
    .join(" ");
}

function isElement(node: Node): node is Element {
  return node.nodeType === Node.ELEMENT_NODE;
}

function serializeChildNode(node: Node): string {
  if (isElement(node)) return serializeXml(node);
  if (
    node.nodeType === Node.TEXT_NODE ||
    node.nodeType === Node.CDATA_SECTION_NODE
  ) {
    return escapeXml(node.nodeValue ?? "");
  }
  if (node.nodeType === Node.COMMENT_NODE) {
    return `<!--${node.nodeValue ?? ""}-->`;
  }
  throw new DocxImportError(
    "unsupported-content",
    "the document holds an XML node we cannot write back out"
  );
}

/**
 * Turns a single element back into an XML string.
 *
 * The browser's XMLSerializer adds fresh declarations such as `xmlns:w` to a fragment it cut out.
 * That declaration already sits at the very top of the document, so we write the string ourselves
 * to keep markup that was not in the original from creeping in.
 */
export function serializeXml(el: Element): string {
  const attrs = attrString(el);
  const open = attrs ? `<${el.nodeName} ${attrs}` : `<${el.nodeName}`;
  const inner = Array.from(el.childNodes).map(serializeChildNode).join("");
  return inner ? `${open}>${inner}</${el.nodeName}>` : `${open}/>`;
}
