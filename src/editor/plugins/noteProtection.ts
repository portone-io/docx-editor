/** Keeps imported footnote and endnote references intact while their bodies are display-only. */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";

function signatures(doc: PMNode): string[] {
  const found: string[] = [];
  doc.descendants((node) => {
    if (node.type.name === "noteReference") {
      found.push(JSON.stringify(node.attrs));
    }
    return true;
  });
  return found;
}

function same(a: readonly string[], b: readonly string[]): boolean {
  return a.length === b.length && a.every((value, index) => value === b[index]);
}

export function noteProtection(): Plugin {
  return new Plugin({
    filterTransaction(transaction, state) {
      return (
        !transaction.docChanged ||
        same(signatures(state.doc), signatures(transaction.doc))
      );
    },
  });
}
