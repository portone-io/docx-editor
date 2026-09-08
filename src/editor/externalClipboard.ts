import {
  type DOMOutputSpec,
  DOMSerializer,
  Fragment,
  type Mark,
  type Node as PMNode,
  Slice,
} from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { styleIdOf } from "../docx/formatting";
import { docxSchema, isPageBreak } from "../schema";
import {
  COPIED_STYLE_ATTRIBUTE,
  readHtmlSlice,
  withPastedContent,
} from "./clipboard/htmlReader";
import { safeHref } from "./clipboard/inlineFormatting";
import { readContextOf } from "./clipboard/readContext";
import { insertPlainText } from "./plainText";
import { moveCaretToDrop } from "./plugins/dropCaret";

/**
 * What a copy is allowed to carry out of the editor.
 *
 * Everything a node or mark draws itself with is private until it is named here: the document's
 * own XML, the identity behind a comment, the body of a comment or a note. A copy lands in
 * whatever application somebody pastes into, so an attribute the schema gains later carries
 * nothing outward until somebody decides it should.
 */
const PUBLIC_ATTRIBUTES: ReadonlySet<string> = new Set([
  "alt",
  "aria-label",
  "class",
  "colspan",
  "height",
  "href",
  "lang",
  "rowspan",
  "src",
  "style",
  "width",
]);

function isAttributes(value: unknown): value is Record<string, unknown> {
  return (
    typeof value === "object" &&
    value !== null &&
    !Array.isArray(value) &&
    !(value instanceof Node)
  );
}

function publicAttributes(
  attrs: Record<string, unknown>
): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(attrs).filter(([name]) => PUBLIC_ATTRIBUTES.has(name))
  );
}

/** A drawing written as a tag name, the attributes it carries, and what stands inside it */
function isSpecArray(value: unknown): value is [string, ...unknown[]] {
  return Array.isArray(value) && typeof value[0] === "string";
}

/** The same drawing with everything the schema did not mean to publish taken off it */
function stripPrivateAttributes(spec: DOMOutputSpec): DOMOutputSpec {
  if (!isSpecArray(spec)) return spec;
  const [tag, ...rest] = spec;
  const attrs = isAttributes(rest[0]) ? publicAttributes(rest[0]) : null;
  const children: unknown[] = (attrs === null ? rest : rest.slice(1)).map(
    (child) => (isSpecArray(child) ? stripPrivateAttributes(child) : child)
  );
  return attrs === null ? [tag, ...children] : [tag, attrs, ...children];
}

function withDomAttribute(
  spec: DOMOutputSpec,
  name: string,
  value: string
): DOMOutputSpec {
  if (!isSpecArray(spec)) return spec;
  const [tag, ...rest] = spec;
  return isAttributes(rest[0])
    ? [tag, { ...rest[0], [name]: value }, ...rest.slice(1)]
    : [tag, { [name]: value }, ...rest];
}

function asTag(spec: DOMOutputSpec, tag: string): DOMOutputSpec {
  return isSpecArray(spec) ? [tag, ...spec.slice(1)] : spec;
}

/**
 * A link written the way everything else writes one.
 *
 * The editor draws the address as data so that a click inside the text places the caret rather
 * than navigating. A copy is read somewhere else, where an anchor is what a link is: it follows
 * in mail or a document, and it comes back as a link when it is pasted here again.
 */
function copiedMarkSpec(mark: Mark, spec: DOMOutputSpec): DOMOutputSpec {
  const stripped = stripPrivateAttributes(spec);
  if (mark.type !== docxSchema.marks.link) return stripped;
  const href = safeHref(
    typeof mark.attrs.href === "string" ? mark.attrs.href : null
  );
  return href === null
    ? stripped
    : asTag(withDomAttribute(stripped, "href", href), "a");
}

function copiedNodeSpec(node: PMNode, spec: DOMOutputSpec): DOMOutputSpec {
  const stripped = stripPrivateAttributes(spec);
  if (node.type !== docxSchema.nodes.paragraph) return stripped;
  const styleId = styleIdOf(node.attrs.pPr);
  return styleId === null
    ? stripped
    : withDomAttribute(stripped, COPIED_STYLE_ATTRIBUTE, styleId);
}

/**
 * The serializer a copy is written with: what the editor draws, with the private attributes taken
 * back off.
 *
 * Wrapping the schema's own drawing rather than declaring a second set of shapes keeps the two
 * from drifting. A node drawn a new way is copied the new way, and only what it publishes changes.
 */
function clipboardSerializer(): DOMSerializer {
  const drawn = DOMSerializer.fromSchema(docxSchema);
  const nodes = Object.fromEntries(
    Object.entries(drawn.nodes).map(([name, toDOM]) => [
      name,
      (node: PMNode) => copiedNodeSpec(node, toDOM(node)),
    ])
  );
  const marks = Object.fromEntries(
    Object.entries(drawn.marks).map(([name, toDOM]) => [
      name,
      (mark: Mark, inline: boolean) =>
        copiedMarkSpec(mark, toDOM(mark, inline)),
    ])
  );
  return new DOMSerializer(nodes, marks);
}

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

export function insertRichHtml(view: EditorView, source: string): boolean {
  const content = readHtmlSlice(
    readContextOf(view.state, view.dom.ownerDocument),
    source
  );
  if (content === null) return false;
  view.dispatch(
    withPastedContent(
      view.state.tr.replaceSelection(content.slice),
      content
    ).scrollIntoView()
  );
  return true;
}

export function insertClipboardData(
  view: EditorView,
  data: { html?: string; text?: string }
): void {
  if (insertRichHtml(view, data.html ?? "")) return;
  insertPlainText(view, data.text ?? "");
}

export function externalClipboard(): Plugin {
  const serializer = clipboardSerializer();
  return new Plugin({
    props: {
      clipboardSerializer: serializer,
      clipboardTextSerializer: clipboardText,
      transformCopied: copiedSlice,
      handlePaste(view, event) {
        insertClipboardData(view, {
          html: event.clipboardData?.getData("text/html"),
          text: event.clipboardData?.getData("text/plain"),
        });
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
