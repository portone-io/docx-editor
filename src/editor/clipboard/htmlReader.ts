import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import {
  type FormattingContext,
  type ParagraphStyleOption,
  paragraphAttrsFor,
} from "../../docx/formatting";
import {
  type ParagraphProps,
  withListNumbering,
  withParagraphStyle,
} from "../../docx/paraProps";
import {
  listsWorn,
  NEW_LISTS_ATTR,
  type NewLists,
  newListsValue,
} from "../../numbering/listRegistry";
import {
  allocateList,
  type ListKind,
  MAX_ILVL,
  templateList,
} from "../../numbering/listTemplate";
import { styleIdOf } from "../../ooxml/props";
import { docxSchema } from "../../schema";
import { COPIED_STYLE_ATTRIBUTE } from "../../schema/clipboard";
import { editorClassNames } from "../../styles/classNames";
import { numIdsIn } from "../commands/listCommands";
import { PASTED_IMAGE_ATTRIBUTE } from "./images";
import {
  appendInline,
  contextFor,
  type InlineContext,
  type InlineStyle,
  marksFor,
  withInlineStyle,
} from "./inlineFormatting";
import type { ListKinds } from "./internalChannel";
import type { HtmlReadContext } from "./readContext";
import { detectHtmlSource } from "./source";

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
  styles: readonly ParagraphStyleOption[],
  sourceStyleId: string | null,
  level: number | null
): string | null {
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
  context: HtmlReadContext,
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
  const styleId = destinationStyleId(
    context.paragraphStyles,
    sourceStyleId,
    level
  );
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
  formatting: FormattingContext,
  list: ListContext | null,
  paragraph: ParagraphProps | null
): Record<string, unknown> | null {
  const props = listParagraphProps(list) ?? paragraph;
  return props
    ? { pPr: props.pPr, ...paragraphAttrsFor(props.pPr, formatting) }
    : null;
}

class HtmlReader {
  readonly blocks: PMNode[] = [];
  readonly used: Set<number>;
  readonly canCreateLists: boolean;
  /**
   * What the pasted paragraphs are resolved against: the document's context, holding the
   * definition of each pasted list as it is read, so that an item is drawn against the list it
   * is joining rather than against a list nothing yet defines.
   */
  private formatting: FormattingContext;

  constructor(
    private readonly context: HtmlReadContext,
    private readonly preserveParagraphStyles: boolean
  ) {
    this.formatting = context.formatting;
    this.used = new Set(context.numbering.used);
    this.canCreateLists = context.numbering.canCreate;
  }

  /** The definitions of every list started while editing, the pasted ones among them */
  get newLists(): NewLists {
    return this.formatting.numbering.added;
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
      const image = token === null ? undefined : this.context.images.get(token);
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
        paragraphAttrs(this.formatting, list, paragraph),
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
          this.context,
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
    const started = allocateList(
      this.formatting.numbering,
      this.used,
      templateList(kind)
    );
    this.formatting = { ...this.formatting, numbering: started.numbering };
    this.used.add(started.numId);
    return started.numId;
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

/**
 * How open the read content is: whether it joins the paragraph it is put into or stands as
 * paragraphs of its own. A copy out of a ProseMirror editor says so itself, and everything else
 * is judged by whether the markup names a block at all.
 */
function sliceDepth(root: ParentNode): 0 | 1 {
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
  /**
   * What the numbers this slice's paragraphs name meant where it was copied
   * (`./internalChannel`). Absent for a reading that gave the lists it read numbers of its own,
   * which knows what kind each is from the markup it read them out of.
   */
  listKinds?: ListKinds;
}

/** The content one piece of already parsed markup reads as, or null when it reads as nothing */
export function readHtml(
  root: ParentNode,
  context: HtmlReadContext
): PastedContent | null {
  const open = sliceDepth(root);
  const reader = new HtmlReader(
    { ...context, source: detectHtmlSource(root) },
    open === 0
  );
  const blocks = reader.read(root);
  return blocks.length === 0
    ? null
    : {
        slice: new Slice(Fragment.fromArray([...blocks]), open, open),
        newLists: reader.newLists,
      };
}

/** The same reading, for a caller holding the markup as a string rather than as elements */
export function readHtmlSlice(
  context: HtmlReadContext,
  source: string
): PastedContent | null {
  if (source.trim() === "") return null;
  const template = context.document.createElement("template");
  template.innerHTML = source;
  return readHtml(template.content, context);
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
