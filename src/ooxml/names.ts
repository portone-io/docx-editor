/**
 * The namespace prefixes the writer spells out, and the namespace each one stands for.
 *
 * Reading is another matter: a document may bind these prefixes to names of its own, so an element
 * that arrived in the file is looked up by its local part. What this package writes is settled
 * here, and the part a fragment is spliced into declares the same binding, so the prefix is
 * decided once rather than at every call site.
 */

import { R_NS, W_NS } from "./xml";

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
