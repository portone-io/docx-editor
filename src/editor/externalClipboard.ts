import {
  type DOMOutputSpec,
  DOMSerializer,
  Fragment,
  type Mark,
  type Node as PMNode,
  Slice,
} from "prosemirror-model";
import { type EditorState, Plugin, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { paragraphAttrsFor, styleIdOf } from "../docx/formatting";
import {
  type ParagraphProps,
  withListNumbering,
  withParagraphStyle,
} from "../docx/paraProps";
import {
  listsWorn,
  NEW_LISTS_ATTR,
  type NewLists,
  newListsValue,
} from "../numbering/listRegistry";
import {
  type ListKind,
  MAX_ILVL,
  nextNumId,
  templateList,
} from "../numbering/listTemplate";
import { docxSchema, isPageBreak } from "../schema";
import { editorClassNames } from "../styles/classNames";
import { PASTED_IMAGE_ATTRIBUTE } from "./clipboard/images";
import {
  appendInline,
  contextFor,
  type InlineContext,
  type InlineStyle,
  marksFor,
  safeHref,
  withInlineStyle,
} from "./clipboard/inlineFormatting";
import { numIdsIn } from "./commands/listCommands";
import { documentFormatting, documentParagraphStyles } from "./documentStyles";
import type { ImageToInsert } from "./insertImage";
import { insertPlainText } from "./plainText";
import { moveCaretToDrop } from "./plugins/dropCaret";
import {
  canStartNewList,
  documentNumbering,
} from "./plugins/numberingDecorations";

const BLOCK_TAGS = new Set([
  "ADDRESS",
  "ARTICLE",
  "ASIDE",
  "BLOCKQUOTE",
  "DIV",
  "H1",
  "H2",
  "H3",
  "H4",
  "H5",
  "H6",
  "LI",
  "P",
  "PRE",
  "SECTION",
]);

const HEADING_LEVELS: Readonly<Record<string, number>> = {
  H1: 1,
  H2: 2,
  H3: 3,
  H4: 4,
  H5: 5,
  H6: 6,
};

const HEADING_FALLBACKS: Readonly<Record<number, InlineStyle>> = {
  1: { bold: true, fontSizePt: 24 },
  2: { bold: true, fontSizePt: 18 },
  3: { bold: true, fontSizePt: 14 },
  4: { bold: true, fontSizePt: 12 },
  5: { bold: true, fontSizePt: 10 },
  6: { bold: true, fontSizePt: 8 },
};

interface ListContext {
  kind: ListKind;
  numId: number | null;
  level: number;
}

interface BlockContext {
  inline: InlineContext;
  paragraph: ParagraphProps | null;
}

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

/**
 * The paragraph style a copy carries, which is the one thing a paste needs that the drawing does
 * not already say. It is the style's id alone, where the editor draws the whole `w:pPr`.
 */
const COPIED_STYLE_ATTRIBUTE = "data-style";

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
const UNREADABLE_BLOCKS: ReadonlySet<string> = new Set([
  "bookmarkBlock",
  "docxRaw",
  "rawBlock",
]);

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

function normalizedStyleName(value: string): string {
  return value.toLowerCase().replace(/[\s_-]+/g, "");
}

function copiedStyleId(element: HTMLElement): string | null {
  if (!element.classList.contains(editorClassNames.paragraph)) return null;
  // A copy taken before the style travelled on its own attribute still carries the whole `w:pPr`,
  // and the id read out of either is checked against this document's styles before it is used
  return (
    element.getAttribute(COPIED_STYLE_ATTRIBUTE) ??
    styleIdOf(element.getAttribute("data-ppr"))
  );
}

function headingLevel(element: HTMLElement, sourceStyleId: string | null) {
  const direct = HEADING_LEVELS[element.tagName];
  if (direct !== undefined) return direct;
  const match = /^heading([1-6])$/.exec(
    normalizedStyleName(sourceStyleId ?? "")
  );
  return match ? Number.parseInt(match[1] ?? "", 10) : null;
}

function destinationStyleId(
  state: EditorState,
  sourceStyleId: string | null,
  level: number | null
): string | null {
  const styles = documentParagraphStyles(state);
  if (sourceStyleId && styles.some((style) => style.id === sourceStyleId)) {
    return sourceStyleId;
  }
  if (level === null) return null;
  const wanted = `heading${level}`;
  return (
    styles.find(
      (style) =>
        normalizedStyleName(style.id) === wanted ||
        normalizedStyleName(style.name) === wanted
    )?.id ?? null
  );
}

function blockContext(
  state: EditorState,
  parent: InlineContext,
  element: HTMLElement,
  preserveParagraphStyle: boolean
): BlockContext {
  const editorParagraph = element.classList.contains(
    editorClassNames.paragraph
  );
  const sourceStyleId = copiedStyleId(element);
  const level = headingLevel(element, sourceStyleId);
  if (!preserveParagraphStyle && editorParagraph) {
    return {
      inline: contextFor(parent, element),
      paragraph: null,
    };
  }
  const styleId = destinationStyleId(state, sourceStyleId, level);
  const paragraph = styleId === null ? null : withParagraphStyle(null, styleId);
  const usesDestinationStyle =
    editorParagraph && (sourceStyleId === null || paragraph !== null);
  let inline = parent;
  if (paragraph === null && level !== null) {
    inline = withInlineStyle(inline, HEADING_FALLBACKS[level] ?? {});
  }
  return {
    inline: contextFor(inline, element, !usesDestinationStyle),
    paragraph,
  };
}

/** A pasted list item joins the list the way a list command puts a paragraph into one */
function listParagraphProps(list: ListContext | null): ParagraphProps | null {
  if (!list || list.numId === null) return null;
  return withListNumbering(null, {
    numbering: { numId: list.numId, ilvl: Math.min(MAX_ILVL, list.level) },
    indent: { kind: "keep" },
  });
}

function paragraphAttrs(
  state: EditorState,
  list: ListContext | null,
  paragraph: ParagraphProps | null
): Record<string, unknown> | null {
  const props = listParagraphProps(list) ?? paragraph;
  return props
    ? {
        pPr: props.pPr,
        ...paragraphAttrsFor(props.pPr, documentFormatting(state)),
      }
    : null;
}

class HtmlReader {
  readonly blocks: PMNode[] = [];
  readonly used: Set<number>;
  readonly canCreateLists: boolean;
  /** The definitions of the lists started while editing, the pasted ones added as they are read */
  private registered: NewLists;

  constructor(
    private readonly state: EditorState,
    private readonly preserveParagraphStyles: boolean,
    private readonly images: ReadonlyMap<string, ImageToInsert>
  ) {
    const numbering = documentNumbering(state);
    this.registered = numbering.added;
    this.used = new Set([
      ...numbering.lists.keys(),
      ...numbering.added.keys(),
      ...numIdsIn(state.doc),
    ]);
    this.canCreateLists = canStartNewList(state);
  }

  /** The definitions of every list started while editing, the pasted ones among them */
  get newLists(): NewLists {
    return this.registered;
  }

  read(root: ParentNode): readonly PMNode[] {
    this.readFlow(root, { style: {}, href: null, preserveWhitespace: false });
    return this.blocks;
  }

  private appendInline(
    target: PMNode[],
    node: Node,
    context: InlineContext
  ): void {
    appendInline(target, node, context, (content, element, inline) => {
      if (element.tagName !== "IMG") return false;
      const token = element.getAttribute(PASTED_IMAGE_ATTRIBUTE);
      const image = token === null ? undefined : this.images.get(token);
      if (!image) return true;
      content.push(
        docxSchema.nodes.image.create(
          {
            ...image,
            // The source drawing may point at a relationship in another package.
            xml: null,
          },
          null,
          marksFor(inline)
        )
      );
      return true;
    });
  }

  private addParagraph(
    content: PMNode[],
    list: ListContext | null = null,
    paragraph: ParagraphProps | null = null
  ): void {
    if (list && list.numId === null) {
      const marker = list.kind === "bullet" ? "• " : "1. ";
      content.unshift(docxSchema.text(marker));
    }
    this.blocks.push(
      docxSchema.nodes.paragraph.create(
        paragraphAttrs(this.state, list, paragraph),
        content
      )
    );
  }

  private readFlow(
    parent: ParentNode,
    context: InlineContext,
    paragraph: ParagraphProps | null = null
  ): void {
    let inline: PMNode[] = [];
    const flush = () => {
      if (inline.length === 0) return;
      this.addParagraph(inline, null, paragraph);
      inline = [];
    };
    for (const child of parent.childNodes) {
      const element =
        child.nodeType === child.ELEMENT_NODE ? (child as HTMLElement) : null;
      if (element?.tagName === "UL" || element?.tagName === "OL") {
        flush();
        this.readList(element, context, 0, null);
      } else if (element && BLOCK_TAGS.has(element.tagName)) {
        flush();
        const before = this.blocks.length;
        const block = blockContext(
          this.state,
          context,
          element,
          this.preserveParagraphStyles
        );
        this.readFlow(element, block.inline, block.paragraph);
        if (this.blocks.length === before) {
          this.addParagraph([], null, block.paragraph);
        }
      } else {
        this.appendInline(inline, child, context);
      }
    }
    flush();
  }

  /**
   * The number a pasted list takes, registered with the definition that list is drawn and written
   * with. A pasted list is started here just as the list button starts one.
   */
  private takeNumId(kind: ListKind): number | null {
    if (!this.canCreateLists) return null;
    const numId = nextNumId(this.used, kind);
    this.registered = new Map([
      ...this.registered,
      [numId, templateList(kind)],
    ]);
    this.used.add(numId);
    return numId;
  }

  private readList(
    list: HTMLElement,
    context: InlineContext,
    level: number,
    inherited: ListContext | null
  ): void {
    const kind: ListKind = list.tagName === "UL" ? "bullet" : "numbered";
    const listContext: ListContext = {
      kind,
      numId: inherited?.kind === kind ? inherited.numId : this.takeNumId(kind),
      level,
    };
    const items = Array.from(list.children).filter(
      (child) => child.tagName === "LI"
    );
    for (const item of items) {
      const itemContext = contextFor(context, item as HTMLElement);
      const content: PMNode[] = [];
      for (const child of item.childNodes) {
        const nested =
          child.nodeType === child.ELEMENT_NODE ? (child as HTMLElement) : null;
        if (nested?.tagName === "UL" || nested?.tagName === "OL") continue;
        this.appendInline(content, child, itemContext);
      }
      this.addParagraph(content, listContext);
      for (const nested of Array.from(item.children)) {
        if (nested.tagName === "UL" || nested.tagName === "OL") {
          this.readList(
            nested as HTMLElement,
            itemContext,
            level + 1,
            listContext
          );
        }
      }
    }
  }
}

function sliceDepth(root: DocumentFragment): 0 | 1 {
  const marker = root.querySelector<HTMLElement>("[data-pm-slice]");
  const openStart = Number.parseInt(
    marker?.getAttribute("data-pm-slice")?.split(/\s+/, 1)[0] ?? "",
    10
  );
  if (Number.isFinite(openStart)) return openStart === 0 ? 0 : 1;
  const blocks = [...BLOCK_TAGS, "OL", "UL"].map((tag) => tag.toLowerCase());
  return root.querySelector(blocks.join(",")) ? 0 : 1;
}

/** What a paste puts in: the content, and the definitions of the lists it started */
export interface PastedContent {
  slice: Slice;
  newLists: NewLists;
}

export function richHtmlSlice(
  state: EditorState,
  document: Document,
  source: string,
  images: ReadonlyMap<string, ImageToInsert> = new Map<string, ImageToInsert>()
): PastedContent | null {
  if (source.trim() === "") return null;
  const template = document.createElement("template");
  template.innerHTML = source;
  const open = sliceDepth(template.content);
  const reader = new HtmlReader(state, open === 0, images);
  const blocks = reader.read(template.content);
  return blocks.length === 0
    ? null
    : {
        slice: new Slice(Fragment.fromArray([...blocks]), open, open),
        newLists: reader.newLists,
      };
}

/**
 * The transaction a paste is put in with: the content, and the definitions of the lists it started
 * recorded on the document node so that the export writes them out and undo takes them back.
 */
export function withPastedContent(
  tr: Transaction,
  content: PastedContent
): Transaction {
  return tr.setDocAttribute(
    NEW_LISTS_ATTR,
    newListsValue(listsWorn(content.newLists, numIdsIn(tr.doc)))
  );
}

export function insertRichHtml(view: EditorView, source: string): boolean {
  const content = richHtmlSlice(view.state, view.dom.ownerDocument, source);
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
