/**
 * The namespace prefixes the writer spells out, and the namespace each one stands for.
 *
 * Reading is another matter: a document may bind these prefixes to names of its own, so an element
 * that arrived in the file is looked up by its local part. What this package writes is settled
 * here, and the part a fragment is spliced into declares the same binding, so the prefix is
 * decided once rather than at every call site.
 */

/** The wordprocessing namespace that every element we read lives in */
export const W_NS =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * The relationship namespace. It is both where the `r:embed` attributes live and the base
 * every relationship type name is built on
 */
export const R_NS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The prefix every WordprocessingML element and attribute this package writes carries */
export const W_PREFIX = "w";

export const NAMESPACES = {
  w: W_NS,
  r: R_NS,
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
} as const;

export type KnownPrefix = keyof typeof NAMESPACES;

/** Whether this is a prefix the package has a namespace for, rather than one a document brought */
export function isKnownPrefix(prefix: string): prefix is KnownPrefix {
  return Object.hasOwn(NAMESPACES, prefix);
}

/** The name `local` goes out under (`qualify("w15", "commentEx")` -> `w15:commentEx`) */
export function qualify(prefix: KnownPrefix, local: string): string {
  return `${prefix}:${local}`;
}

/** `wName("val")` -> `w:val`. The one place the written WordprocessingML prefix is decided */
export function wName(local: string): string {
  return qualify(W_PREFIX, local);
}

/** The declaration a part root carries for one prefix, ready to write inside an opening tag */
export function xmlnsDecl(prefix: KnownPrefix): string {
  return `xmlns:${prefix}="${NAMESPACES[prefix]}"`;
}
