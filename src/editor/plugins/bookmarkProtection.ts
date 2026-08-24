/** Prevents ordinary editor transactions from dropping imported bookmark range markers. */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";

const BOOKMARK_XML = /<(?:[\w.-]+:)?bookmark(?:Start|End)\b/;

function signatures(doc: PMNode): string[] {
  const found: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "bookmarkBlock") {
      found.push(`block:${node.attrs.srcId}:${node.attrs.name}`);
    } else if (
      node.type.name === "rawInline" &&
      typeof node.attrs.xml === "string" &&
      BOOKMARK_XML.test(node.attrs.xml)
    ) {
      found.push(`inline:${node.attrs.xml}`);
    }
    return true;
  });
  return found;
}

function same(a: readonly string[], b: readonly string[]): boolean {
  return a.length === b.length && a.every((value, index) => value === b[index]);
}

export function bookmarkProtection(): Plugin {
  return new Plugin({
    filterTransaction(transaction, state) {
      return (
        !transaction.docChanged ||
        same(signatures(state.doc), signatures(transaction.doc))
      );
    },
  });
}
