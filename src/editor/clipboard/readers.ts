import { Fragment, type ResolvedPos, Slice } from "prosemirror-model";
import { docxSchema } from "../../schema";
import { toPieces } from "../plainText";
import { type PastedContent, readHtml } from "./htmlReader";
import type { HtmlReadContext } from "./readContext";

/** What one reader is handed: the markup on the clipboard, and the document it is read into */
export interface ClipboardInput {
  dom: Node;
  context: HtmlReadContext;
}

/**
 * One way of reading a clipboard. The chain is tried in order and the first reader to answer
 * decides what goes in, so a reader that does not recognize what it was handed answers nothing.
 */
export type ClipboardReader = (input: ClipboardInput) => PastedContent | null;

/** Everything the editor did not write itself, read for the formatting the markup states */
export const foreignHtmlReader: ClipboardReader = ({ dom, context }) =>
  dom instanceof Element || dom instanceof DocumentFragment
    ? readHtml(dom, context)
    : null;

export const DEFAULT_READERS: readonly ClipboardReader[] = [foreignHtmlReader];

/**
 * Plain text as content of this document: docx has no character for a line break or a tab stop,
 * so those become the elements it does have, and the rest keeps the formatting in force where it
 * is put.
 */
export function plainTextSlice(text: string, $context: ResolvedPos): Slice {
  const marks = $context.marks();
  const nodes = toPieces(text).map((piece) => {
    if ("text" in piece) return docxSchema.text(piece.text, marks);
    if (piece.node === "hardBreak") {
      return docxSchema.nodes.hardBreak.create(null, null, marks);
    }
    const beside = marks.filter((mark) => mark.type !== docxSchema.marks.tab);
    return docxSchema.text(
      "\t",
      docxSchema.marks.tab.create().addToSet(beside)
    );
  });
  return nodes.length === 0
    ? Slice.empty
    : new Slice(Fragment.fromArray(nodes), 0, 0);
}
