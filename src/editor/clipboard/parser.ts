import { DOMParser, Slice } from "prosemirror-model";
import { docxSchema } from "../../schema";
import type { PastedContent } from "./htmlReader";
import type { HtmlReadContext } from "./readContext";
import type { ClipboardReader } from "./readers";

/**
 * The parser ProseMirror hands the clipboard to.
 *
 * It is a `DOMParser` because that is what the `clipboardParser` prop is, but it holds no rules:
 * clipboard markup is read by the reader chain alone, so a `parseDOM` rule of the schema - which
 * would take the document's own XML back out of a private attribute - never sees it.
 */
export class DocxClipboardParser extends DOMParser {
  private read: PastedContent | null = null;

  constructor(
    private readonly readers: readonly ClipboardReader[],
    private readonly contextOf: () => HtmlReadContext | null
  ) {
    super(docxSchema, []);
  }

  parseSlice(dom: Node): Slice {
    this.read = null;
    const context = this.contextOf();
    if (context === null) return Slice.empty;
    for (const reader of this.readers) {
      const content = reader({ dom, context });
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
}
