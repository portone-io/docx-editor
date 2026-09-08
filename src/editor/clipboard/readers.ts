import { Fragment, type ResolvedPos, Slice } from "prosemirror-model";
import { NO_NEW_LISTS } from "../../numbering/listRegistry";
import { docxSchema } from "../../schema";
import { toPieces } from "../plainText";
import { type PastedContent, readHtml } from "./htmlReader";
import { recallCopied } from "./internalChannel";
import type { HtmlReadContext } from "./readContext";

/** What one reader is handed: the markup on the clipboard, and the document it is read into */
export interface ClipboardInput {
  dom: Node;
  /** The name the markup carries for a slice this editor kept (`./internalChannel`). Null when it carries none */
  token: string | null;
  /** The open document the markup is being read into. Null for a state built without one */
  sessionId: string | null;
  context: HtmlReadContext;
}

/**
 * One way of reading a clipboard. The chain is tried in order and the first reader to answer
 * decides what goes in, so a reader that does not recognize what it was handed answers nothing.
 */
export type ClipboardReader = (input: ClipboardInput) => PastedContent | null;

/**
 * A copy made in this same open document, handed back as the slice it was.
 *
 * It leads the chain because it is the only reader that loses nothing: everything below reads
 * markup, and markup is all a copy from anywhere else amounts to here.
 */
export const internalSliceReader: ClipboardReader = ({ token, sessionId }) => {
  const slice = recallCopied(token, sessionId);
  return slice === null ? null : { slice, newLists: NO_NEW_LISTS };
};

/** Everything the editor did not write itself, read for the formatting the markup states */
export const foreignHtmlReader: ClipboardReader = ({ dom, context }) =>
  dom instanceof Element || dom instanceof DocumentFragment
    ? readHtml(dom, context)
    : null;

export const DEFAULT_READERS: readonly ClipboardReader[] = [
  internalSliceReader,
  foreignHtmlReader,
];

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
