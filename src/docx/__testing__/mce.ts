/**
 * Test-only MCE preprocessing for the namespaces the validation schema set understands.
 * ECMA-376 Part 3, sections 7 and 9: resolve declarations in their original scope, discard only
 * unknown ignorable markup, and validate the branch selected by the consumer's configuration.
 * Preservation hints and application-defined extension elements are outside this harness's
 * profile. Unsupported MCE directives fail explicitly instead of hiding content from xmllint.
 */
import { readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { parseXml } from "../../ooxml/xml";

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const XSD_NS = "http://www.w3.org/2001/XMLSchema";

/** Follow the same schema imports as entry.xsd, rather than claiming to understand all OOXML. */
function schemaNamespaces(): ReadonlySet<string> {
  const namespaces = new Set([XML_NS]);
  const visited = new Set<string>();
  function visit(path: string): void {
    if (visited.has(path)) return;
    visited.add(path);
    const root = parseXml(readFileSync(path, "utf8")).documentElement;
    const namespace = root.getAttribute("targetNamespace");
    if (namespace) namespaces.add(namespace);
    for (const child of Array.from(root.children)) {
      if (
        child.namespaceURI !== XSD_NS ||
        !["import", "include"].includes(child.localName)
      )
        continue;
      const location = child.getAttribute("schemaLocation");
      if (location) visit(resolve(dirname(path), location));
    }
  }
  visit(
    resolve(
      dirname(fileURLToPath(import.meta.url)),
      "../../../spec/schemas/transitional/wml.xsd"
    )
  );
  return namespaces;
}

const UNDERSTOOD = schemaNamespaces();
interface Context {
  readonly ignorable: ReadonlySet<string>;
  readonly processContent: readonly (readonly [string, string])[];
}
const EMPTY: Context = { ignorable: new Set(), processContent: [] };

function tokens(value: string): string[] {
  return value.trim() === "" ? [] : value.trim().split(/\s+/);
}

function namespaceOf(el: Element, prefix: string): string {
  const namespace = el.lookupNamespaceURI(prefix);
  if (namespace === null || namespace === MC_NS) {
    throw new Error(`Invalid MCE namespace prefix: ${prefix}`);
  }
  return namespace;
}

function contextAt(el: Element, inherited: Context): Context {
  for (const attr of Array.from(el.attributes)) {
    if (
      attr.namespaceURI === MC_NS &&
      !["Ignorable", "ProcessContent", "MustUnderstand"].includes(
        attr.localName
      )
    ) {
      throw new Error(`Unsupported MCE directive: ${attr.localName}`);
    }
  }
  const ignorable = new Set(inherited.ignorable);
  for (const prefix of tokens(el.getAttributeNS(MC_NS, "Ignorable") ?? "")) {
    ignorable.add(namespaceOf(el, prefix));
  }
  if (el.namespaceURI === MC_NS) {
    for (const attr of Array.from(el.attributes)) {
      const namespace = attr.namespaceURI;
      if (namespace === "http://www.w3.org/2000/xmlns/" || namespace === MC_NS)
        continue;
      if (
        namespace === null &&
        el.localName === "Choice" &&
        attr.localName === "Requires"
      )
        continue;
      if (
        namespace !== null &&
        namespace !== XML_NS &&
        ignorable.has(namespace)
      )
        continue;
      throw new Error(`Invalid attribute on ${el.localName}: ${attr.name}`);
    }
  }
  const processContent = [...inherited.processContent];
  for (const token of tokens(
    el.getAttributeNS(MC_NS, "ProcessContent") ?? ""
  )) {
    const parts = token.split(":");
    if (parts.length !== 2 || !parts[0] || !parts[1])
      throw new Error(`Invalid ProcessContent name: ${token}`);
    const namespace = namespaceOf(el, parts[0]);
    if (!ignorable.has(namespace))
      throw new Error(`ProcessContent namespace is not ignorable: ${token}`);
    processContent.push([namespace, parts[1]]);
  }
  return { ignorable, processContent };
}

function mustUnderstand(el: Element): void {
  for (const prefix of tokens(
    el.getAttributeNS(MC_NS, "MustUnderstand") ?? ""
  )) {
    if (!UNDERSTOOD.has(namespaceOf(el, prefix)))
      throw new Error(`Unsupported MustUnderstand namespace: ${prefix}`);
  }
}

function ignored(el: Element, context: Context): boolean {
  return (
    el.namespaceURI !== null &&
    context.ignorable.has(el.namespaceURI) &&
    !UNDERSTOOD.has(el.namespaceURI)
  );
}

/** Process in the original tree before unwrapping: both prefix bindings and MCE scope live there. */
function unwrap(el: Element, context: Context): void {
  if (
    ["base", "lang", "space"].some((name) => el.hasAttributeNS(XML_NS, name))
  ) {
    throw new Error("ProcessContent cannot unwrap XML context attributes");
  }
  rewriteInside(el, context);
  const parent = el.parentNode;
  if (parent === null) throw new Error("MCE element has no parent");
  for (const child of Array.from(el.childNodes)) parent.insertBefore(child, el);
  el.remove();
}

function alternate(el: Element, context: Context): void {
  mustUnderstand(el);
  let selected: Element | null = null;
  let selectedContext = context;
  let choices = 0;
  let fallbackSeen = false;
  for (const node of Array.from(el.childNodes)) {
    if (node.nodeType === 3 && node.textContent?.trim())
      throw new Error("Text inside AlternateContent");
  }
  for (const child of Array.from(el.children)) {
    const childContext = contextAt(child, context);
    if (ignored(child, childContext)) continue;
    if (child.namespaceURI !== MC_NS)
      throw new Error("Unexpected AlternateContent child");
    for (const attr of Array.from(child.attributes)) {
      if (
        attr.namespaceURI === null &&
        !(child.localName === "Choice" && attr.localName === "Requires")
      ) {
        throw new Error("Unexpected alternate-content branch attribute");
      }
    }
    if (child.localName === "Choice") {
      if (fallbackSeen) throw new Error("Choice follows Fallback");
      choices += 1;
      const requires = tokens(child.getAttribute("Requires") ?? "");
      if (requires.length === 0)
        throw new Error("Choice has no Requires namespaces");
      const namespaces = requires.map((prefix) => namespaceOf(child, prefix));
      if (
        selected === null &&
        namespaces.every((namespace) => UNDERSTOOD.has(namespace))
      ) {
        selected = child;
        selectedContext = childContext;
      }
    } else if (child.localName === "Fallback") {
      if (fallbackSeen || choices === 0)
        throw new Error("Fallback must follow Choices and occur at most once");
      fallbackSeen = true;
      if (selected === null) {
        selected = child;
        selectedContext = childContext;
      }
    } else {
      throw new Error("Unexpected MCE element inside AlternateContent");
    }
  }
  if (choices === 0) throw new Error("AlternateContent has no Choice");
  if (selected !== null) {
    mustUnderstand(selected);
    rewriteInside(selected, selectedContext);
    const parent = el.parentNode;
    if (parent === null) throw new Error("AlternateContent has no parent");
    for (const child of Array.from(selected.childNodes))
      parent.insertBefore(child, el);
  }
  el.remove();
}

function rewriteInside(el: Element, context: Context): void {
  for (const child of Array.from(el.children)) rewrite(child, context);
}

function rewrite(el: Element, inherited: Context): void {
  const context = contextAt(el, inherited);
  if (ignored(el, context)) {
    const process = context.processContent.some(
      ([namespace, name]) =>
        namespace === el.namespaceURI && (name === "*" || name === el.localName)
    );
    if (process) {
      mustUnderstand(el);
      unwrap(el, context);
    } else el.remove();
    return;
  }
  if (el.namespaceURI === MC_NS) {
    if (el.localName !== "AlternateContent")
      throw new Error(`Unexpected MCE element: ${el.localName}`);
    alternate(el, context);
    return;
  }
  mustUnderstand(el);
  rewriteInside(el, context);
  for (const attr of Array.from(el.attributes)) {
    const namespace = attr.namespaceURI;
    if (
      namespace === MC_NS ||
      (namespace !== null &&
        context.ignorable.has(namespace) &&
        !UNDERSTOOD.has(namespace))
    ) {
      el.removeAttributeNode(attr);
    }
  }
}

export function withoutIgnorableMarkup(xml: string): string {
  const doc = parseXml(xml);
  rewrite(doc.documentElement, EMPTY);
  return new XMLSerializer().serializeToString(doc);
}
