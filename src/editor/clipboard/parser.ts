import { DOMParser, Slice } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { docxSchema } from "../../schema";
import type { PastedContent } from "./htmlReader";
import type { ClipboardInput, ClipboardReader } from "./readers";

/**
 * The parser ProseMirror hands the clipboard to.
 *
 * It is a `DOMParser` because that is what the `clipboardParser` prop is, but it holds no rules:
 * clipboard markup is read by the reader chain alone, so a `parseDOM` rule of the schema - which
 * would take the document's own XML back out of a private attribute - never sees it.
 */
export class DocxClipboardParser extends DOMParser {
  private read: PastedContent | null = null;
  private textOnly = false;
  private pasteEvent: ClipboardEvent | null = null;

  constructor(
    private readonly readers: readonly ClipboardReader[],
    private readonly inputOf: (dom: Node) => ClipboardInput | null
  ) {
    super(docxSchema, []);
  }

  parseSlice(dom: Node): Slice {
    this.read = null;
    this.setPlainText(false);
    const input = this.inputOf(dom);
    if (input === null) return Slice.empty;
    for (const reader of this.readers) {
      const content = reader(input);
      if (content === null) continue;
      this.read = content;
      return content.slice;
    }
    return Slice.empty;
  }

  /**
   * What the last parse read, given up as it is taken.
   *
   * ProseMirror widens whatever a parser returns to the openness the surrounding markup implies
   * and wraps it in the nodes a `data-pm-slice` context names, neither of which the reader asked
   * for. The plugin puts this back in `transformPasted`, and takes it so that a later drag - which
   * reaches `transformPasted` without a parse - never picks up a stale reading.
   */
  takeRead(): PastedContent | null {
    const read = this.read;
    this.read = null;
    return read;
  }

  /** A new reading replaces the mode left by the previous clipboard event */
  setPlainText(plain: boolean): void {
    this.textOnly = plain;
    this.pasteEvent = null;
  }

  isPlainText(event: ClipboardEvent): boolean {
    // A file-only paste skips parsing altogether. Bind a reading to the first event its handlers
    // see, so a later event cannot reuse the text-only decision from an earlier paste.
    if (this.pasteEvent === null) this.pasteEvent = event;
    return this.pasteEvent === event && this.textOnly;
  }
}

/** Whether the current paste chose text, so image handlers must leave its HTML and files alone */
export function plainTextPaste(
  view: EditorView,
  event: ClipboardEvent
): boolean {
  const parser = view.someProp("clipboardParser");
  return parser instanceof DocxClipboardParser && parser.isPlainText(event);
}
