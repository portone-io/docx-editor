import type { Node as PMNode } from "prosemirror-model";
import {
  type EditorState,
  Plugin,
  PluginKey,
  type Transaction,
} from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { editorClassNames, editorCssVariables } from "../../styles/classNames";

interface TabDecorationState {
  widths: ReadonlyMap<number, number>;
  decorations: DecorationSet;
}

const tabDecorationsKey = new PluginKey<TabDecorationState>(
  "docxEditorTabDecorations"
);

function decorationsFor(
  doc: PMNode,
  widths: ReadonlyMap<number, number>
): DecorationSet {
  const decorations: Decoration[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText || !node.text?.includes("\t")) return true;
    for (let offset = 0; offset < node.text.length; offset += 1) {
      if (node.text[offset] !== "\t") continue;
      const from = pos + offset;
      const width = widths.get(from);
      const attrs: Record<string, string> = {
        class: editorClassNames.tabSlot,
        // Distinct attributes prevent adjacent ranges from sharing one DOM wrapper.
        "data-tab-position": `${from}`,
      };
      if (width !== undefined) {
        attrs["data-tab-layout"] = "";
        attrs.style = `${editorCssVariables.tabWidth}:${Math.max(0, width)}px`;
      }
      decorations.push(
        Decoration.inline(from, from + 1, attrs, {
          inclusiveStart: false,
          inclusiveEnd: false,
        })
      );
    }
    return true;
  });
  return DecorationSet.create(doc, decorations);
}

function characterAt(doc: PMNode, position: number): string {
  return doc.textBetween(position, position + 1, "", "");
}

function mappedWidths(
  transaction: Transaction,
  widths: ReadonlyMap<number, number>
): Map<number, number> {
  const mapped = new Map<number, number>();
  for (const [position, width] of widths) {
    const result = transaction.mapping.mapResult(position, 1);
    if (
      !result.deleted &&
      result.pos < transaction.doc.content.size &&
      characterAt(transaction.doc, result.pos) === "\t"
    ) {
      mapped.set(result.pos, width);
    }
  }
  return mapped;
}

export function tabWidths(state: EditorState): ReadonlyMap<number, number> {
  return tabDecorationsKey.getState(state)?.widths ?? new Map();
}

export function setTabWidths(
  transaction: Transaction,
  widths: ReadonlyMap<number, number>
): Transaction {
  return transaction.setMeta(tabDecorationsKey, widths);
}

/** Gives every tab character an independent DOM range without changing the document model. */
export function tabDecorations(): Plugin<TabDecorationState> {
  return new Plugin<TabDecorationState>({
    key: tabDecorationsKey,
    state: {
      init: (_config, state) => {
        const widths = new Map<number, number>();
        return { widths, decorations: decorationsFor(state.doc, widths) };
      },
      apply(transaction, current, _oldState, newState) {
        const supplied: unknown = transaction.getMeta(tabDecorationsKey);
        if (!transaction.docChanged && !(supplied instanceof Map)) {
          return current;
        }
        const widths =
          supplied instanceof Map
            ? (supplied as ReadonlyMap<number, number>)
            : mappedWidths(transaction, current.widths);
        return {
          widths,
          decorations: decorationsFor(newState.doc, widths),
        };
      },
    },
    props: {
      decorations: (state) =>
        tabDecorationsKey.getState(state)?.decorations ?? null,
    },
  });
}
