import type { EditorState } from "prosemirror-state";
import type {
  FormattingContext,
  ParagraphStyleOption,
} from "../../docx/formatting";
import { numIdsIn } from "../commands/listCommands";
import { documentFormatting, documentParagraphStyles } from "../documentStyles";
import type { ImageToInsert } from "../insertImage";
import { canStartNewList } from "../plugins/numberingDecorations";

/**
 * Everything a piece of HTML is read against.
 *
 * The reader answers about the document it is read into - which styles it may name, which list
 * numbers are taken, whether a list may be started at all - and nothing about the editor holding
 * it. Gathering those in one place is what keeps the reader off the editing state, so the same
 * reading can be asked for from a paste, from a drop, and from a test.
 */
export interface HtmlReadContext {
  /** The document the markup is read into an element of */
  document: Document;
  formatting: FormattingContext;
  paragraphStyles: readonly ParagraphStyleOption[];
  numbering: {
    used: ReadonlySet<number>;
    canCreate: boolean;
  };
  /** The images already loaded for this read, by the token the markup carries */
  images: ReadonlyMap<string, ImageToInsert>;
}

export function readContextOf(
  state: EditorState,
  document: Document,
  images: ReadonlyMap<string, ImageToInsert> = new Map<string, ImageToInsert>()
): HtmlReadContext {
  return {
    document,
    formatting: documentFormatting(state),
    paragraphStyles: documentParagraphStyles(state),
    numbering: {
      used: numIdsIn(state.doc),
      canCreate: canStartNewList(state),
    },
    images,
  };
}
