/**
 * Creates, removes, and inspects link spans. Existing links are edited as whole objects, while
 * unlocked plain-text pieces can be linked independently; non-inclusive edges remain outside.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { docxSchema } from "../../schema";
import { rangeTouchesLocked } from "../../schema/locks";

const linkType = docxSchema.marks.link;

function text(value: unknown): string | null {
  return typeof value === "string" ? value : null;
}

/** An address naming a scheme of its own, which is written as it reads (RFC 3986 §3.1) */
const NAMES_SCHEME = /^[a-z][a-z0-9+.-]*:/i;

/** An address meaning a place rather than a host: a file beside the document, or a bookmark */
const MEANS_A_PLACE = /^[./#]/;

/**
 * The address as it goes into the document.
 *
 * A relationship target with no scheme is read as relative to the package, so a host typed bare
 * would name a file beside the document rather than a site. Word and Google Docs both settle that
 * as the link is made, and that is the end to settle it at: the file then holds an address that
 * says what it means, wherever it is opened. `http` is what they settle on and what this follows.
 */
function settled(address: string): string {
  const typed = address.trim();
  if (NAMES_SCHEME.test(typed) || MEANS_A_PLACE.test(typed)) return typed;
  const host = typed.split(/[/?#]/, 1)[0];
  return host.includes(".") ? `http://${typed}` : typed;
}

function linkMarkOf(marks: readonly Mark[]): Mark | null {
  return marks.find((mark) => mark.type === linkType) ?? null;
}

/** One piece of inline content a link can go on, and the link it already wears */
interface LinkPiece {
  from: number;
  to: number;
  link: Mark | null;
}

/**
 * Whether a link may go on this inline node.
 *
 * The two are asked apart because a text node's own type allows no marks at all in prosemirror's
 * model: what governs the marks on text is the textblock holding it, which is also the node an
 * `addMark` step asks.
 */
function linkable(node: PMNode, parent: PMNode | null): boolean {
  if (parent === null || !parent.type.allowsMarkType(linkType)) return false;
  return node.isText || node.type.allowsMarkType(linkType);
}

/**
 * The pieces of inline content inside this stretch.
 *
 * A piece per inline node, so that a lock shutting one of them leaves the others open, the way
 * character formatting works. Content no link may go on is skipped, and a whole paragraph or table
 * the stretch runs across is walked into.
 */
function piecesIn(doc: PMNode, from: number, to: number): LinkPiece[] {
  const pieces: LinkPiece[] = [];
  doc.nodesBetween(from, to, (node, pos, parent) => {
    if (!node.isInline) return true;
    if (!linkable(node, parent)) return false;
    const start = Math.max(pos, from);
    const end = Math.min(pos + node.nodeSize, to);
    if (start < end) {
      pieces.push({ from: start, to: end, link: linkMarkOf(node.marks) });
    }
    return false;
  });
  return pieces;
}

/** One link inside a textblock, as the one whole stretch it covers */
interface LinkSpan {
  from: number;
  to: number;
  link: Mark;
}

/**
 * Every link in this textblock, each as the one whole stretch it covers.
 *
 * A link that wrapped several runs comes in as several inlines wearing the very same mark, and what
 * an edit to a link is about is the link rather than the run standing inside it.
 */
function linkSpans(parent: PMNode, start: number): LinkSpan[] {
  const spans: LinkSpan[] = [];
  parent.forEach((child, offset) => {
    const link = linkMarkOf(child.marks);
    if (!link) return;
    const from = start + offset;
    const to = from + child.nodeSize;
    const last = spans.at(-1);
    if (last && last.to === from && last.link.eq(link)) last.to = to;
    else spans.push({ from, to, link });
  });
  return spans;
}

/** The link the caret stands inside, as the whole stretch it covers. null where it stands in none */
function caretSpan(state: EditorState): LinkSpan | null {
  const $from = state.selection.$from;
  const link = linkMarkOf($from.marks());
  if (!link) return null;
  return (
    linkSpans($from.parent, $from.start()).find(
      (span) =>
        span.link.eq(link) && span.from <= $from.pos && $from.pos <= span.to
    ) ?? null
  );
}

/**
 * The link a selection sits inside, as the whole stretch it covers. null where it runs past either
 * edge of a link, where it covers two of them, and where the block it stands in holds no text.
 *
 * The edges count here, unlike for a caret: a selection covering a link exactly has not left it.
 */
function selectionSpan(state: EditorState): LinkSpan | null {
  const { $from, $to } = state.selection;
  if (!$from.parent.isTextblock || $from.parent !== $to.parent) return null;
  return (
    linkSpans($from.parent, $from.start()).find(
      (span) => span.from <= $from.pos && $to.pos <= span.to
    ) ?? null
  );
}

/**
 * Every link this stretch touches, each as the whole stretch it covers.
 *
 * A stretch ending exactly where a link begins touches nothing: that is the same edge rule the
 * caret reads, and it is what leaves a selection up to a link alone.
 */
function linkSpansTouching(doc: PMNode, from: number, to: number): LinkSpan[] {
  const spans: LinkSpan[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return true;
    for (const span of linkSpans(node, pos + 1)) {
      if (span.from < to && from < span.to) spans.push(span);
    }
    return false;
  });
  return spans;
}

/**
 * What a link edit reaches: the whole of every link the selection touches, and the pieces of the
 * selection standing in no link.
 *
 * A caret reaches the one link it stands in, and a caret standing in no link reaches nothing.
 * Putting the address in as text of its own is not something this editor does, so there is nothing
 * there to wrap.
 *
 * A whole link is one piece here rather than one per inline, which is what makes a lock over any
 * part of it drop the link entire: editing a link up to a lock would leave the text in three, link,
 * plain and link, which is the split an edit to a link must never make.
 */
function openPieces(state: EditorState): LinkPiece[] {
  const found = state.selection.empty
    ? caretPieces(state)
    : selectionPieces(state);
  return found.filter(
    (piece) => !rangeTouchesLocked(state.doc, piece.from, piece.to)
  );
}

function caretPieces(state: EditorState): LinkPiece[] {
  const span = caretSpan(state);
  return span === null ? [] : [span];
}

function selectionPieces(state: EditorState): LinkPiece[] {
  return state.selection.ranges.flatMap((range) => [
    ...linkSpansTouching(state.doc, range.$from.pos, range.$to.pos),
    ...piecesIn(state.doc, range.$from.pos, range.$to.pos).filter(
      (piece) => piece.link === null
    ),
  ]);
}

/**
 * The mark this piece is to wear.
 *
 * A piece already inside a link keeps everything that link carries and only changes where it
 * points, so its tooltip, its history flag and the rest of its opening tag survive the retargeting
 * (`docx/hyperlink` rewrites the one attribute). A piece inside none gets a link of its own, which
 * the export writes an opening tag and a relationship for.
 */
function linkMarkFor(piece: LinkPiece, href: string): Mark {
  return linkType.create({ ...piece.link?.attrs, href });
}

/**
 * Puts a link at this address on the text the selection covers, and changes where every link it
 * touches points, each of them whole.
 *
 * A piece already pointing there is left untouched, so the wrapper it came in with is not rewritten
 * for nothing. A host typed with no scheme is written with one (`settled`).
 */
export function setLink(url: string): Command {
  return (state, dispatch) => {
    const href = settled(url);
    if (href === "") return false;
    const pieces = openPieces(state).filter(
      (piece) => text(piece.link?.attrs.href) !== href
    );
    if (pieces.length === 0) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const piece of pieces) {
        tr.addMark(piece.from, piece.to, linkMarkFor(piece, href));
      }
      dispatch(tr);
    }
    return true;
  };
}

/**
 * Takes off every link the selection touches, each of them whole, however little of it the
 * selection covers.
 * The text and its formatting stay where they are; only the wrapper goes.
 */
export const removeLink: Command = (state, dispatch) => {
  const pieces = openPieces(state).filter((piece) => piece.link !== null);
  if (pieces.length === 0) return false;
  if (dispatch) {
    const tr = state.tr;
    for (const piece of pieces) {
      tr.removeMark(piece.from, piece.to, linkType);
    }
    dispatch(tr);
  }
  return true;
};

/**
 * Whether a link can go on where the selection stands, which is what a control offering one is
 * drawn from.
 *
 * It is asked before an address has been typed, which is why it is a query of its own rather than
 * `setLink(url)` with no dispatch.
 */
export function canSetLink(state: EditorState): boolean {
  return openPieces(state).length > 0;
}

/**
 * The address at the selection. null where the selection is in no link, where the link it is in
 * names a bookmark rather than an address, and where several addresses are mixed.
 *
 * The three answer as one, unlike `activeFontSize`, because the one thing drawn from this is the
 * address field of the link panel and all three leave it empty. It reads the same pieces the
 * commands edit, so the address shown is the address `setLink` would change.
 */
export function activeLink(state: EditorState): string | null {
  const pieces = openPieces(state);
  const first = text(pieces[0]?.link?.attrs.href);
  return pieces.every((piece) => text(piece.link?.attrs.href) === first)
    ? first
    : null;
}

/** The stretch one link covers, and where it points */
export interface ActiveLinkSpan {
  from: number;
  to: number;
  /** The address. null for a link that names a bookmark rather than an address */
  href: string | null;
}

/**
 * The one link the selection sits inside, as the stretch it covers and the address it points at.
 * null where the selection is in no link, where it runs past the edge of one, and where it covers
 * several.
 *
 * This is what the card that says where a link points is drawn from (`ui/LinkCard`), so it answers
 * about the link rather than about the address: a link naming a bookmark is here with a null `href`,
 * where `activeLink` and it are one answer.
 *
 * The locks are the one thing it does not read. A link inside a settled part of the document still
 * has an address worth reading, and what may be done to it is `canSetLink` and `removeLink`.
 */
export function activeLinkSpan(state: EditorState): ActiveLinkSpan | null {
  const span = state.selection.empty ? caretSpan(state) : selectionSpan(state);
  if (!span) return null;
  return { from: span.from, to: span.to, href: text(span.link.attrs.href) };
}
