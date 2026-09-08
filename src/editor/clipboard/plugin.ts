/**
 * The one plugin holding the editor's clipboard.
 *
 * Everything the clipboard carries in or out passes through the props declared here: what a copy
 * is written as, what a paste is read as, and what a drop from outside becomes. ProseMirror asks
 * these before it asks any `handlePaste`. This plugin handles only an empty reading, preserving
 * the selection or retrying the text fallback. Content goes in through ProseMirror's insertion
 * with the `paste` and `uiEvent` meta a plugin
 * watching for one reads (`prosemirror-view`'s `doPaste`), and a table decides for itself what a
 * paste over a cell selection means (`prosemirror-tables`).
 */

import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  NEW_LISTS_ATTR,
  type NewLists,
  NO_NEW_LISTS,
  newListsOf,
  newListsValue,
} from "../../numbering/listRegistry";
import { docxSchema, isPageBreak } from "../../schema";
import { clipboardSerializer } from "../../schema/clipboard";
import { numIdsIn } from "../commands/listCommands";
import { documentOf } from "../editorDocument";
import { insertPlainText } from "../plainText";
import { moveCaretToDrop } from "../plugins/dropCaret";
import { documentNumbering } from "../plugins/numberingDecorations";
import {
  type CopyRoute,
  INTERNAL_TOKEN_ATTRIBUTE,
  internalTokenOf,
  rememberCopied,
} from "./internalChannel";
import {
  DEFAULT_NORMALIZERS,
  normalizePasted,
  type SliceNormalizer,
} from "./normalizers";
import { DocxClipboardParser } from "./parser";
import { readContextOf } from "./readContext";
import {
  type ClipboardReader,
  DEFAULT_READERS,
  plainTextSlice,
} from "./readers";

/** The blocks that stand for something the editor never read, which read as nothing at all */
const UNREADABLE_BLOCKS: ReadonlySet<string> = new Set(["rawBlock"]);

function stringAttr(value: unknown): string {
  return typeof value === "string" ? value : "";
}

/**
 * What one inline node says when the copy is read as text.
 *
 * A comment marker and a preserved fragment say nothing: they stand for something around the
 * text rather than in it. An image says what it was given to say instead.
 */
function inlineText(node: PMNode): string {
  // A tab is written as the character it is, so a run of them says how many there were. The mark
  // beside it is what the editor draws the stop with, not what the text says
  if (node.isText) return node.text ?? "";
  if (node.type === docxSchema.nodes.hardBreak) {
    return isPageBreak(node.attrs.brAttrs) ? "\f" : "\n";
  }
  if (node.type === docxSchema.nodes.image) return stringAttr(node.attrs.alt);
  if (node.type === docxSchema.nodes.noteReference) {
    // A note whose own mark follows draws no number, and the text reads as the page does
    return node.attrs.customMarkFollows === true
      ? ""
      : stringAttr(node.attrs.label);
  }
  return "";
}

function rowsText(table: PMNode): string {
  const rows: string[] = [];
  table.forEach((row) => {
    const cells: string[] = [];
    row.forEach((cell) => {
      cells.push(fragmentText(cell.content));
    });
    rows.push(cells.join("\t"));
  });
  return rows.join("\n");
}

function blockText(node: PMNode): string {
  if (node.type === docxSchema.nodes.table) return rowsText(node);
  if (UNREADABLE_BLOCKS.has(node.type.name)) return "";
  return fragmentText(node.content);
}

/**
 * A fragment read as text. Its children are all inline or all block, which is what decides
 * whether they run together or stand on lines of their own.
 */
function fragmentText(fragment: Fragment): string {
  const pieces: string[] = [];
  fragment.forEach((child) => {
    pieces.push(child.isInline ? inlineText(child) : blockText(child));
  });
  return pieces.join(fragment.firstChild?.isInline === true ? "" : "\n");
}

/**
 * The wrappers a copy is open through, emptied of what they carry.
 *
 * `prosemirror-view` writes those wrappers into `data-pm-slice` as JSON once the serializer has
 * run, so a table's, a row's and a cell's own XML would leave the editor there whatever the
 * drawing says. It peels a wrapper only while the slice is open past it on both sides and it holds
 * a single child, so those are the ones emptied here; a paste reads the open depth, which the
 * shape still says, and a drag inside the editor carries the slice itself rather than the text.
 */
function bareWrappers(
  content: Fragment,
  openStart: number,
  openEnd: number
): Fragment {
  const wrapper = content.firstChild;
  if (openStart <= 1 || openEnd <= 1 || content.childCount !== 1)
    return content;
  if (wrapper === null || wrapper.childCount !== 1) return content;
  return Fragment.from(
    wrapper.type.create(
      null,
      bareWrappers(wrapper.content, openStart - 1, openEnd - 1),
      wrapper.marks
    )
  );
}

function copiedSlice(slice: Slice): Slice {
  return new Slice(
    bareWrappers(slice.content, slice.openStart, slice.openEnd),
    slice.openStart,
    slice.openEnd
  );
}

/**
 * The plain text a copy leaves beside the HTML.
 *
 * Prosemirror's own answer is the text content with a line between blocks, which loses a tab, a
 * page break, and the difference between the next cell and the next row. Somewhere those are the
 * whole of what was copied, a table pasted into a spreadsheet above all.
 */
function clipboardText(slice: Slice): string {
  return fragmentText(slice.content);
}

/**
 * The modifier that turns a drag inside the editor into a copy, which `prosemirror-view` reads off
 * the platform the same way (`dragMoves` in its input handling).
 */
function dragCopyModifier(): "altKey" | "ctrlKey" {
  return /Mac|iP(hone|[oa]d)/.test(navigator.platform) ? "altKey" : "ctrlKey";
}

/**
 * Whether this drop moves what it carries rather than leaving a copy of it behind.
 *
 * ProseMirror settles that at the drop and not at the dragstart, so `view.dragging.move` is only
 * what the drag set out as: the modifier held down over the drop overrules it. It asks after the
 * slice has been through `transformPasted`, which is why the same question is asked here, off the
 * same `dragCopies` prop and the same modifier, rather than read back from `handleDrop`.
 */
function dropMoves(view: EditorView, event: DragEvent): boolean {
  if (view.dragging === null) return false;
  let copies: boolean | undefined;
  view.someProp("dragCopies", (test) => {
    copies = copies === true || test(event);
  });
  return copies === undefined ? !event[dragCopyModifier()] : !copies;
}

export interface ClipboardOptions {
  /** Tried from the front. The first reader to answer decides what a paste puts in */
  readers?: readonly ClipboardReader[];
  /** Run in order over every pasted and dropped slice before it is put in */
  normalizers?: readonly SliceNormalizer[];
}

export function docxClipboard(options: ClipboardOptions = {}): Plugin {
  const normalizers = options.normalizers ?? DEFAULT_NORMALIZERS;
  /** The name given to the copy being written, which the serializer runs straight after */
  let copyToken: string | null = null;
  const serializer = clipboardSerializer((copy) => {
    if (copyToken !== null) {
      copy.firstElementChild?.setAttribute(INTERNAL_TOKEN_ATTRIBUTE, copyToken);
    }
  });
  let host: EditorView | null = null;
  const parser = new DocxClipboardParser(
    options.readers ?? DEFAULT_READERS,
    (dom) =>
      host === null
        ? null
        : {
            dom,
            token: internalTokenOf(dom),
            sessionId: documentOf(host.state).session?.sessionId ?? null,
            context: readContextOf(host.state, host.dom.ownerDocument),
          }
  );
  /** The definitions the last reading started, held until the edit carrying them lands */
  let started: NewLists | null = null;
  /** Whether the drop being handled moves what it carries rather than copying it */
  let dropMove = false;
  /** Whether the serialization being written is the one the dragstart being handled asked for */
  let draggingOut = false;

  return new Plugin({
    view(view) {
      host = view;
      return {
        destroy() {
          if (host === view) host = null;
        },
      };
    },
    /**
     * The definitions of the lists a paste began, registered on the document once the paste is
     * in. A list the pasted markup started is numbered as it is read, and only the edit that
     * lands says whether that number is worn in the end.
     */
    appendTransaction(transactions, _before, after) {
      const lists = started;
      started = null;
      if (
        lists === null ||
        lists.size === 0 ||
        !transactions.some((tr) => tr.docChanged)
      ) {
        return null;
      }
      const worn = numIdsIn(after.doc);
      const registered = newListsOf(after.doc.attrs[NEW_LISTS_ATTR]);
      const missing = [...lists].filter(
        ([numId]) => worn.has(numId) && !registered.has(numId)
      );
      if (missing.length === 0) return null;
      return after.tr.setDocAttribute(
        NEW_LISTS_ATTR,
        newListsValue(new Map([...registered, ...missing]))
      );
    },
    props: {
      handleDOMEvents: {
        paste() {
          // A file-only clipboard skips both parsers. Clear the previous mode even if a consumer
          // handled that paste before its event reached our built-in image handlers.
          parser.setPlainText(false);
          return false;
        },
        // ProseMirror writes and reads the dragged slice inside the very event the two below
        // answer for, through props it hands the view and not the event, and it runs these first.
        // A microtask is the first thing to run once the event is over, so neither answer is left
        // standing for whatever comes next
        dragstart() {
          draggingOut = true;
          queueMicrotask(() => {
            draggingOut = false;
          });
          return false;
        },
        drop(view, event) {
          dropMove = dropMoves(view, event);
          queueMicrotask(() => {
            dropMove = false;
          });
          return false;
        },
      },
      clipboardParser: parser,
      clipboardSerializer: serializer,
      clipboardTextParser(text, context, plain) {
        // `plain` here means Shift/API text paste. transformPasted's similarly named argument
        // is also true for an ordinary paste with no HTML, which may still carry an image file.
        parser.setPlainText(plain);
        return plainTextSlice(text, context);
      },
      clipboardTextSerializer: clipboardText,
      transformCopied(slice, view) {
        // The slice is kept as it stands: what the wrappers are emptied of below leaves for the
        // clipboard, and a paste back into this session is given what was copied instead
        const route: CopyRoute = draggingOut ? "drag" : "clipboard";
        copyToken = rememberCopied(
          route,
          slice,
          documentOf(view.state).session?.sessionId ?? null,
          documentNumbering(view.state)
        );
        return copiedSlice(slice);
      },
      transformPasted(slice, view, plain) {
        if (!plain) parser.setPlainText(false);
        const read = parser.takeRead();
        const content = normalizePasted(
          read ?? { slice, newLists: NO_NEW_LISTS },
          view.state,
          dropMove,
          normalizers
        );
        started = content.newLists;
        return content.slice;
      },
      handlePaste(view, event, slice) {
        if (slice.content.size > 0) return false;
        // No reader recognized the HTML. Its text may still be useful; an empty text reading,
        // including stripped control characters, must not turn a paste into a deletion.
        const text = event.clipboardData?.getData("text/plain") ?? "";
        if (!parser.isPlainText(event) && text) view.pasteText(text, event);
        return true;
      },
      handleDrop(view, event) {
        if (view.dragging) return false;
        const text = event.dataTransfer?.getData("text/plain") ?? "";
        if (text) {
          moveCaretToDrop(view, event);
          insertPlainText(view, text);
        }
        return true;
      },
    },
  });
}
