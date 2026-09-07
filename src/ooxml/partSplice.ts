/**
 * Putting children into the root element of a package part, and leaving every byte around them as
 * the part arrived with.
 *
 * A part is text the writers cut and join, never a tree they build again, so the spots to cut at
 * are found in the text: the root's opening tag past the prolog, and its closing tag counted to
 * by depth. Searching for the closing tag instead would take a `</Relationships>` inside a
 * comment for the end of the part, and miss the root altogether when it is written under a
 * prefix or closes on itself.
 */

import { childOrderOf } from "./childOrder";
import { attrsText, type XmlAttr } from "./element";
import { DocxExportError } from "./errors";
import type { KnownPrefix } from "./names";
import { parseAttrs, readTag, rootTagAt, type Tag } from "./tagScan";
import { localPart } from "./xml";

/** One child to put into a part, named by the local part of its tag so the child order can place it */
export interface PartChild {
  name: string;
  xml: string;
}

/**
 * What goes into a part's root. Every field is optional and they compose: the children are
 * replaced first, then the inserts are placed among them, then the prepend and the append go on
 * either end, and last the opening tag is rewritten.
 */
export interface PartSplice {
  /** The root's local name. Any prefix, a self-closing root, a prolog and comments are all handled */
  root: string;
  /** Children put at the spot `CHILD_ORDER[root]` lays down, ahead of the first child that order puts after them */
  insert?: readonly PartChild[];
  /** Put in as the last children */
  append?: string;
  /** Put in right after the opening tag */
  prepend?: string;
  /** Written in place of every child */
  replaceChildren?: string;
  /**
   * Rewrites the opening tag, handed as it is about to be written: opened when children go into a
   * root that closed on itself, else as it arrived.
   */
  rootTag?: (open: string) => string;
}

interface PartRoot {
  openAt: number;
  open: Tag;
  /** Where the closing tag starts. null for a root that closes on itself */
  closeAt: number | null;
}

/** Where the root's closing tag starts, counting depth from just inside its opening tag. null when it never closes */
function closingTagAt(xml: string, open: Tag): number | null {
  let depth = 1;
  let at = open.end;
  for (;;) {
    const lt = xml.indexOf("<", at);
    if (lt === -1) return null;
    const tag = readTag(xml, lt);
    if (tag === null) return null;
    if (tag.kind === "open") depth += 1;
    if (tag.kind === "close") {
      depth -= 1;
      if (depth === 0) return tag.name === open.name ? lt : null;
    }
    at = tag.end;
  }
}

function readRoot(
  xml: string,
  root: string
): { root: PartRoot } | { problem: string } {
  const openAt = rootTagAt(xml);
  const open = openAt === -1 ? null : readTag(xml, openAt);
  if (open === null || open.kind === "close" || localPart(open.name) !== root) {
    return { problem: `the ${root} part has no ${root} root element` };
  }
  if (open.kind === "empty") return { root: { openAt, open, closeAt: null } };
  const closeAt = closingTagAt(xml, open);
  if (closeAt === null) {
    return { problem: `the ${root} part has no closing ${root} tag` };
  }
  return { root: { openAt, open, closeAt } };
}

/**
 * Why the part cannot be spliced around its root, or null when it can. The same refusal
 * `splicePart` throws, so a caller may ask ahead and be answered in the same words.
 */
export function partRootProblem(xml: string, root: string): string | null {
  const reading = readRoot(xml, root);
  return "problem" in reading ? reading.problem : null;
}

/** The direct children of a root, each as its local name and where its tag starts in `inner` */
function directChildren(inner: string): { name: string; at: number }[] {
  const children: { name: string; at: number }[] = [];
  let depth = 0;
  let at = 0;
  for (;;) {
    const lt = inner.indexOf("<", at);
    if (lt === -1) return children;
    const tag = readTag(inner, lt);
    if (tag === null) return children;
    if (depth === 0 && (tag.kind === "open" || tag.kind === "empty")) {
      children.push({ name: localPart(tag.name), at: lt });
    }
    if (tag.kind === "open") depth += 1;
    if (tag.kind === "close") depth -= 1;
    at = tag.end;
  }
}

/**
 * The spot for a child of this name: ahead of the first child the order puts after it, on the
 * end when none does. A child already there that the order does not know decides nothing.
 * Throws for a name the order does not know at all, the way `props` does, since writing it on the
 * end would hide a missing registry entry behind an export a validator may refuse.
 */
function insertionAt(
  children: readonly { name: string; at: number }[],
  name: string,
  order: readonly string[],
  root: string,
  end: number
): number {
  const target = order.indexOf(name);
  if (target === -1) {
    throw new Error(
      `${name} is not a child the order of ${root} knows; add it to CHILD_ORDER`
    );
  }
  const after = children.find((child) => order.indexOf(child.name) > target);
  return after === undefined ? end : after.at;
}

function withInserted(
  inner: string,
  inserts: readonly PartChild[],
  root: string
): string {
  if (inserts.length === 0) return inner;
  const order = childOrderOf(root);
  const children = directChildren(inner);
  const placed = inserts
    .map((child) => ({
      at: insertionAt(children, child.name, order, root, inner.length),
      xml: child.xml,
      rank: order.indexOf(child.name),
    }))
    .sort((a, b) => a.at - b.at || a.rank - b.rank);
  let out = "";
  let from = 0;
  for (const { at, xml } of placed) {
    out += inner.slice(from, at) + xml;
    from = at;
  }
  return out + inner.slice(from);
}

/**
 * The part with the children of `splice` put into its root, and every byte outside what was put
 * in standing as it arrived. A root that closed on itself is opened to hold them.
 *
 * Throws `malformed-xml` naming the part when the root is not there or never closes.
 */
export function splicePart(xml: string, splice: PartSplice): string {
  const reading = readRoot(xml, splice.root);
  if ("problem" in reading) {
    throw new DocxExportError("malformed-xml", reading.problem);
  }
  const { openAt, open, closeAt } = reading.root;
  const inner = closeAt === null ? "" : xml.slice(open.end, closeAt);
  const children =
    (splice.prepend ?? "") +
    withInserted(
      splice.replaceChildren ?? inner,
      splice.insert ?? [],
      splice.root
    ) +
    (splice.append ?? "");

  const openTag = xml.slice(openAt, open.end);
  const opened =
    closeAt === null && children !== "" ? `${openTag.slice(0, -2)}>` : openTag;
  const head = xml.slice(0, openAt) + (splice.rootTag?.(opened) ?? opened);
  if (closeAt !== null) return head + children + xml.slice(closeAt);
  if (children === "") return head + xml.slice(open.end);
  return `${head}${children}</${open.name}>${xml.slice(open.end)}`;
}

/** The name of the root element exactly as the part writes it, prefix included. null when the part holds none */
function rootNameOf(xml: string): string | null {
  const at = rootTagAt(xml);
  const open = at === -1 ? null : readTag(xml, at);
  return open === null || open.kind === "close" ? null : open.name;
}

/**
 * The prefix the root's name carries, colon included, so a child written beside the ones the part
 * holds is spelled the way the root is. Empty for a root written under the default namespace, and
 * for a part with no root, which the splice that follows refuses.
 */
export function rootPrefixOf(xml: string): string {
  const name = rootNameOf(xml) ?? "";
  const colon = name.indexOf(":");
  return colon === -1 ? "" : name.slice(0, colon + 1);
}

export interface RootDeclarations {
  /** The prefixes the root has to bind, and what to bind each to where it does not yet */
  namespaces: Partial<Record<KnownPrefix, string>>;
  /** The tokens `mc:Ignorable` has to name, added to the ones it names already */
  ignorable?: readonly string[];
}

const IGNORABLE = /\smc:Ignorable\s*=\s*(["'])([^"']*)\1/;

/**
 * The opening tag with the declarations it lacks written on the end of it. What it declares
 * already stands as it was written, quoting and spacing included.
 */
function withDeclarations(
  openTag: string,
  { namespaces, ignorable }: RootDeclarations
): string {
  const tag = readTag(openTag, 0);
  const closeLength = tag?.kind === "empty" ? 2 : 1;
  const attrs =
    tag === null
      ? null
      : parseAttrs(openTag.slice(tag.nameEnd, tag.end - closeLength));
  if (tag === null || attrs === null) {
    throw new DocxExportError(
      "malformed-xml",
      "the root element's attributes could not be read"
    );
  }
  const declared = new Set(attrs.map(([name]) => name));
  const additions: XmlAttr[] = [];
  for (const [prefix, namespace] of Object.entries(namespaces)) {
    if (namespace !== undefined && !declared.has(`xmlns:${prefix}`)) {
      additions.push([`xmlns:${prefix}`, namespace]);
    }
  }
  const ignoring = IGNORABLE.exec(openTag);
  let tokens: string[] | null = null;
  if (ignorable !== undefined && ignorable.length > 0) {
    if (ignoring === null) {
      additions.push(["mc:Ignorable", ignorable.join(" ")]);
    } else {
      const had = (attrs.find(([name]) => name === "mc:Ignorable")?.[1] ?? "")
        .split(/\s+/)
        .filter(Boolean);
      const missing = ignorable.filter((token) => !had.includes(token));
      if (missing.length > 0) tokens = [...had, ...missing];
    }
  }
  let next = openTag;
  if (ignoring !== null && tokens !== null) {
    const rewritten = ` mc:Ignorable=${ignoring[1]}${tokens.join(" ")}${ignoring[1]}`;
    next =
      next.slice(0, ignoring.index) +
      rewritten +
      next.slice(ignoring.index + ignoring[0].length);
  }
  if (additions.length === 0) return next;
  const end = next.length - closeLength;
  return `${next.slice(0, end)} ${attrsText(additions)}${next.slice(end)}`;
}

/**
 * The part with its root binding every prefix named and its `mc:Ignorable` naming every token
 * asked for, each declared once: one the root declares already is left as it was written.
 */
export function ensureRootDeclarations(
  xml: string,
  declarations: RootDeclarations
): string {
  const root = rootNameOf(xml);
  if (root === null) {
    throw new DocxExportError(
      "malformed-xml",
      "the part has no root element to declare a namespace on"
    );
  }
  return splicePart(xml, {
    root: localPart(root),
    rootTag: (open) => withDeclarations(open, declarations),
  });
}
