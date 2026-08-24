import { AllSelection, type Selection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { docxSchema } from "../schema";

/** A single character that becomes an element, or text that goes in as it is */
type Piece = { text: string } | { node: "hardBreak" | "tab" };

/**
 * Control characters that become elements because they do not exist as characters in docx.
 * The vertical tab (U+000B) and the form feed (U+000C) are the line breaks Word writes to the
 * clipboard.
 */
const AS_NODE: Record<string, "hardBreak" | "tab"> = {
  "\n": "hardBreak",
  "\u000B": "hardBreak",
  "\u000C": "hardBreak",
  "\t": "tab",
};

// biome-ignore lint/suspicious/noControlCharactersInRegex: sorting out control characters is what this file does
const SPLIT_AT = /([\n\u000B\u000C\t])/;

/** Control characters we strip because they carry no meaning and XML 1.0 cannot hold them */
// biome-ignore lint/suspicious/noControlCharactersInRegex: same as above
const FORBIDDEN = /[\u0000-\u0008\u000E-\u001F]/g;

/**
 * Splits the incoming text into the pieces to be put into the document.
 * Control characters that carry meaning become elements, and the rest are stripped because they
 * would make the file impossible to open.
 */
function toPieces(source: string): Piece[] {
  return source
    .replace(/\r\n?/g, "\n")
    .replace(FORBIDDEN, "")
    .split(SPLIT_AT)
    .flatMap((chunk): Piece[] => {
      const node = AS_NODE[chunk];
      if (node) return [{ node }];
      return chunk ? [{ text: chunk }] : [];
    });
}

/**
 * Inserts text at the caret position.
 * Since docx has no such character for a line break, breaks go in as break elements (`w:br`)
 * and tabs as tab elements (`w:tab`).
 */
function insertWithSelection(
  view: EditorView,
  source: string,
  selection: Selection
): void {
  const pieces = toPieces(source);
  if (pieces.length === 0) return;

  const tr = view.state.tr.setSelection(selection);
  for (const piece of pieces) {
    if ("text" in piece) tr.insertText(piece.text);
    else if (piece.node === "hardBreak") {
      tr.replaceSelectionWith(docxSchema.nodes.hardBreak.create());
    } else {
      const marks = (tr.storedMarks ?? tr.selection.$from.marks()).filter(
        (mark) => mark.type !== docxSchema.marks.tab
      );
      tr.replaceSelectionWith(
        docxSchema.text("\t", docxSchema.marks.tab.create().addToSet(marks)),
        false
      );
    }
  }
  view.dispatch(tr.scrollIntoView());
}

export function insertPlainText(view: EditorView, source: string): void {
  insertWithSelection(view, source, view.state.selection);
}

export function insertPlainTextAt(
  view: EditorView,
  source: string,
  from: number,
  to: number
): void {
  const { doc } = view.state;
  const selection =
    from === 0 && to === doc.content.size
      ? new AllSelection(doc)
      : TextSelection.between(doc.resolve(from), doc.resolve(to));
  insertWithSelection(view, source, selection);
}
