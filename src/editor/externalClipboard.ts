import { Fragment, type Node as PMNode, Slice } from "prosemirror-model";
import { type EditorState, Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { effectiveParagraphFormat, styleIdOf } from "../docx/formatting";
import { type ParagraphProps, withParagraphStyle } from "../docx/paraProps";
import { type ListKind, MAX_ILVL, nextNumId } from "../numbering/listTemplate";
import { docxSchema } from "../schema";
import { editorClassNames } from "../styles/classNames";
import { PASTED_IMAGE_ATTRIBUTE } from "./clipboard/images";
import {
  appendInline,
  contextFor,
  type InlineContext,
  type InlineStyle,
  marksFor,
  withInlineStyle,
} from "./clipboard/inlineFormatting";
import { listRefOf } from "./commands/listCommands";
import {
  defaultParagraphStyleId,
  documentParagraphFormatting,
  documentParagraphStyles,
  documentStyles,
} from "./documentStyles";
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

function normalizedStyleName(value: string): string {
  return value.toLowerCase().replace(/[\s_-]+/g, "");
}

function copiedStyleId(element: HTMLElement): string | null {
  return element.classList.contains(editorClassNames.paragraph)
    ? styleIdOf(element.getAttribute("data-ppr"))
    : null;
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
  const paragraph =
    styleId === null
      ? null
      : withParagraphStyle(
          null,
          styleId,
          documentStyles(state),
          defaultParagraphStyleId(state)
        );
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

function listParagraphAttrs(
  state: EditorState,
  list: ListContext | null
): Record<string, unknown> | null {
  if (!list || list.numId === null) return null;
  const ilvl = Math.min(MAX_ILVL, list.level);
  const pPr =
    `<w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/>` +
    `<w:numId w:val="${list.numId}"/></w:numPr></w:pPr>`;
  return {
    pPr,
    format: effectiveParagraphFormat(pPr, documentParagraphFormatting(state)),
  };
}

function paragraphAttrs(
  state: EditorState,
  list: ListContext | null,
  paragraph: ParagraphProps | null
): Record<string, unknown> | null {
  const listed = listParagraphAttrs(state, list);
  if (listed) return listed;
  return paragraph
    ? {
        pPr: paragraph.pPr,
        format: effectiveParagraphFormat(
          paragraph.pPr,
          documentParagraphFormatting(state)
        ),
        styleRun: paragraph.styleRun,
      }
    : null;
}

function usedNumIds(state: EditorState): Set<number> {
  const used = new Set(documentNumbering(state).lists.keys());
  state.doc.descendants((node) => {
    if (node.type !== docxSchema.nodes.paragraph) return true;
    const ref = listRefOf(node);
    if (ref) used.add(ref.numId);
    return false;
  });
  return used;
}

class HtmlReader {
  readonly blocks: PMNode[] = [];
  readonly used: Set<number>;
  readonly canCreateLists: boolean;

  constructor(
    private readonly state: EditorState,
    private readonly preserveParagraphStyles: boolean,
    private readonly images: ReadonlyMap<string, ImageToInsert>
  ) {
    this.used = usedNumIds(state);
    this.canCreateLists = canStartNewList(state);
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

  private takeNumId(kind: ListKind): number | null {
    if (!this.canCreateLists) return null;
    const id = nextNumId(this.used, kind);
    this.used.add(id);
    return id;
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

export function richHtmlSlice(
  state: EditorState,
  document: Document,
  source: string,
  images: ReadonlyMap<string, ImageToInsert> = new Map<string, ImageToInsert>()
): Slice | null {
  if (source.trim() === "") return null;
  const template = document.createElement("template");
  template.innerHTML = source;
  const open = sliceDepth(template.content);
  const blocks = new HtmlReader(state, open === 0, images).read(
    template.content
  );
  return blocks.length === 0
    ? null
    : new Slice(Fragment.fromArray([...blocks]), open, open);
}

export function insertRichHtml(view: EditorView, source: string): boolean {
  const slice = richHtmlSlice(view.state, view.dom.ownerDocument, source);
  if (slice === null) return false;
  view.dispatch(view.state.tr.replaceSelection(slice).scrollIntoView());
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
  return new Plugin({
    props: {
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
