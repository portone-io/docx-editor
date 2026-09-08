/**
 * What each node and mark looks like once it has left the editor.
 *
 * A node draws itself twice. `toDOM` draws the page, where the editor reads its own attributes
 * back (an IME composition, a browser edit), so everything the document is written from stands on
 * the element. A copy lands in Word, in a spreadsheet, in a mail: none of them can read a `w:pPr`,
 * and none of them should be handed a comment's author or the body of a footnote. So the shape a
 * copy takes is declared here, beside the schema, rather than worked out afterwards by taking the
 * page's drawing apart.
 *
 * The two cannot drift apart unnoticed: `Readonly<Record<DocxNodeName, ...>>` refuses to compile
 * until a node added to the schema says what it leaves as, and `clipboard.test.ts` asks the schema
 * itself the same question at runtime.
 */

import {
  type DOMOutputSpec,
  DOMSerializer,
  Fragment,
  type Mark,
  type Node as PMNode,
  type Schema,
  type Slice,
} from "prosemirror-model";
import {
  spanCount,
  toCellFormat,
  toGridCols,
  toParagraphFormat,
  toRowFormat,
  toRunFormat,
  toTableFormat,
  toTableWidth,
} from "../model/format";
import { emuToPx, toImageExtent, toImageSrc } from "../ooxml/image";
import { styleIdOf } from "../ooxml/props";
import { editorClassNames } from "../styles/classNames";
import { DEFAULT_FONT_FALLBACKS } from "../styles/fontStack";
import {
  cellStyle,
  columnWidthPx,
  paragraphStyle,
  rowStyle,
  runStyle,
  tableStyle,
} from "../styles/inlineStyle";
import { type docxSchema, isPageBreak } from "./docxSchema";

/**
 * The paragraph style a copy carries, which is the one thing a paste needs that the drawing does
 * not already say. It is the style's id alone, where the editor draws the whole `w:pPr`.
 */
export const COPIED_STYLE_ATTRIBUTE = "data-style";

/**
 * The address a link may carry, whichever way it travels.
 *
 * A copy writes one out and a paste reads one in, and an address this rule turns down is one
 * neither end should act on, so both ask here.
 */
export function safeHref(value: string | null): string | null {
  const href = value?.trim() ?? "";
  if (href === "") return null;
  if (/^(?:https?|mailto|tel):/i.test(href)) return href;
  return /^(?:[./]|#)/.test(href) ? href : null;
}

export interface ClipboardNodeSpec {
  /**
   * The HTML a copy carries. `null` in place of a function says this kind of node never leaves the
   * editor; a function answering null says this one leaves nothing. Either way neither the node
   * nor anything inside it is written.
   */
  toClipboardDOM: ((node: PMNode) => DOMOutputSpec | null) | null;
  /**
   * The plain text a copy carries, given what its children said. A null answer leaves the node out
   * of the text entirely, separator and all, so a block standing for something invisible does not
   * open a blank line where it stood.
   */
  toClipboardText:
    | ((node: PMNode, children: readonly string[]) => string | null)
    | null;
}

export interface ClipboardMarkSpec {
  /** null writes no wrapper, and the content it covered goes out on its own */
  toClipboardDOM: ((mark: Mark) => DOMOutputSpec) | null;
}

/**
 * `Schema.nodes` is indexed by `string` as well as by the names it was built from, so `keyof` over
 * it widens to `string`. The generic argument the schema literal was inferred with keeps them.
 */
export type DocxNodeName =
  typeof docxSchema extends Schema<infer N, string> ? N : never;
export type DocxMarkName =
  typeof docxSchema extends Schema<string, infer M> ? M : never;

function textAttr(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

/** The number a note reference draws, and nothing where the note draws a mark of its own */
function noteLabel(node: PMNode): string {
  return node.attrs.customMarkFollows === true
    ? ""
    : (textAttr(node.attrs.label) ?? "");
}

function spanAttribute(count: unknown): string | undefined {
  const span = spanCount(count);
  return span > 1 ? `${span}` : undefined;
}

/**
 * What a fragment the editor kept rather than modelled leaves as.
 *
 * Only what it draws travels. A `w:cr` is a line and a `w:noBreakHyphen` is a character, so both
 * are text wherever they are pasted; a chip is the editor naming the element that stands here,
 * which is scaffolding rather than content and means nothing in another application.
 */
function preservedText(node: PMNode): string | null {
  const display = node.attrs.display;
  if (display === "break") return "\n";
  if (display === "text") return textAttr(node.attrs.text) ?? "";
  return null;
}

/** A preserved fragment leaves the one way, whichever level it was kept at */
const PRESERVED_SPEC: ClipboardNodeSpec = {
  toClipboardDOM: (node) => {
    const said = preservedText(node);
    if (said === null) return null;
    return said === "\n" ? ["br"] : ["span", said];
  },
  toClipboardText: preservedText,
};

/** Content that stands between the text rather than in it, which no other application can hold */
const WRITES_NOTHING: ClipboardNodeSpec = {
  toClipboardDOM: null,
  toClipboardText: null,
};

const CLIPBOARD_NODES: Readonly<Record<DocxNodeName, ClipboardNodeSpec>> = {
  doc: {
    toClipboardDOM: null,
    toClipboardText: (_node, children) => children.join("\n"),
  },
  paragraph: {
    toClipboardDOM: (node) => [
      "p",
      {
        class: editorClassNames.paragraph,
        style: paragraphStyle(
          toParagraphFormat(node.attrs.format),
          toRunFormat(node.attrs.styleRun)
        ),
        [COPIED_STYLE_ATTRIBUTE]: styleIdOf(node.attrs.pPr) ?? undefined,
      },
      0,
    ],
    toClipboardText: (_node, children) => children.join(""),
  },
  table: {
    toClipboardDOM: (node) => {
      const attrs = {
        class: editorClassNames.table,
        style: tableStyle(
          toTableFormat(node.attrs.format),
          toTableWidth(node.attrs.tblW)
        ),
      };
      const body = ["tbody", 0];
      const columns = toGridCols(node.attrs.gridCols).map((dxa) => [
        "col",
        { style: `width:${columnWidthPx(dxa)}px` },
      ]);
      if (columns.length === 0) return ["table", attrs, body];
      return ["table", attrs, ["colgroup", ...columns], body];
    },
    toClipboardText: (_node, children) => children.join("\n"),
  },
  tableRow: {
    toClipboardDOM: (node) => [
      "tr",
      {
        class: editorClassNames.tableRow,
        style: rowStyle(toRowFormat(node.attrs.format)),
      },
      0,
    ],
    toClipboardText: (_node, children) => children.join("\t"),
  },
  tableCell: {
    toClipboardDOM: (node) => [
      "td",
      {
        class: editorClassNames.tableCell,
        style: cellStyle(toCellFormat(node.attrs.format)),
        colspan: spanAttribute(node.attrs.colspan),
        rowspan: spanAttribute(node.attrs.rowspan),
      },
      0,
    ],
    // A tab stands between two cells and a line between two rows wherever a table is pasted as
    // text, so a cell of several paragraphs says them on one line rather than breaking its row
    toClipboardText: (_node, children) => children.join(" "),
  },
  rawBlock: WRITES_NOTHING,
  hardBreak: {
    toClipboardDOM: () => ["br"],
    // A page break ends more than a line, and the form feed is what says so in plain text
    toClipboardText: (node) => (isPageBreak(node.attrs.brAttrs) ? "\f" : "\n"),
  },
  image: {
    toClipboardDOM: (node) => {
      const extent = toImageExtent(node.attrs.extent);
      return [
        "img",
        {
          class: editorClassNames.image,
          src: toImageSrc(node.attrs.src) ?? undefined,
          alt: textAttr(node.attrs.alt) ?? "",
          // The measure the document keeps is its own. The pixels it was drawn at are what another
          // application reads a size from, and what a paste back into this editor reads one from
          width: extent ? `${Math.round(emuToPx(extent.cx))}` : undefined,
          height: extent ? `${Math.round(emuToPx(extent.cy))}` : undefined,
        },
      ];
    },
    // A picture is not text. But a copy may be read where no picture can follow, and there the
    // words the document gave it for that very case are better than a gap
    toClipboardText: (node) => textAttr(node.attrs.alt) ?? "",
  },
  commentStart: WRITES_NOTHING,
  commentEnd: WRITES_NOTHING,
  commentReference: WRITES_NOTHING,
  noteReference: {
    // A note whose own mark follows draws no number, and an empty superscript would stand for
    // nothing wherever the copy lands
    toClipboardDOM: (node) => {
      const label = noteLabel(node);
      return label === "" ? null : ["sup", label];
    },
    toClipboardText: noteLabel,
  },
  rawRunContent: PRESERVED_SPEC,
  rawInline: PRESERVED_SPEC,
  text: {
    // `DOMSerializer` writes a text node itself and never asks a spec for one, so there is no
    // drawing to declare. The characters travel as they stand
    toClipboardDOM: null,
    toClipboardText: (node) => node.text ?? "",
  },
};

const CLIPBOARD_MARKS: Readonly<Record<DocxMarkName, ClipboardMarkSpec>> = {
  sdt: { toClipboardDOM: () => ["span", 0] },
  link: {
    // The editor draws the address as data so that a click inside the text places the caret.
    // Anywhere else an anchor is what a link is: it follows in mail or a document, and it comes
    // back as a link when it is pasted here again
    toClipboardDOM: (mark) => {
      const href = safeHref(textAttr(mark.attrs.href) ?? null);
      const attrs = { class: editorClassNames.link, href: href ?? undefined };
      return href === null ? ["span", attrs, 0] : ["a", attrs, 0];
    },
  },
  run: {
    toClipboardDOM: (mark) => {
      const format = toRunFormat(mark.attrs.format);
      return [
        "span",
        {
          class: editorClassNames.run,
          style: runStyle(format, DEFAULT_FONT_FALLBACKS),
          // Which shape of a Han character is drawn is decided by this and nothing else
          lang: format?.lang,
        },
        0,
      ];
    },
  },
  tab: { toClipboardDOM: () => ["span", { class: editorClassNames.tab }, 0] },
};

export const clipboardSpecs: {
  nodes: Readonly<Record<DocxNodeName, ClipboardNodeSpec>>;
  marks: Readonly<Record<DocxMarkName, ClipboardMarkSpec>>;
} = { nodes: CLIPBOARD_NODES, marks: CLIPBOARD_MARKS };

function nodeSpec(node: PMNode): ClipboardNodeSpec | undefined {
  return CLIPBOARD_NODES[node.type.name as DocxNodeName];
}

function clipboardDOMOf(node: PMNode): DOMOutputSpec | null {
  const spec = nodeSpec(node);
  if (spec === undefined || spec.toClipboardDOM === null) return null;
  return spec.toClipboardDOM(node);
}

/** The children of a fragment that leave anything at all */
function written(fragment: Fragment): Fragment {
  const kept: PMNode[] = [];
  fragment.forEach((child) => {
    if (child.isText || clipboardDOMOf(child) !== null) kept.push(child);
  });
  return kept.length === fragment.childCount ? fragment : Fragment.from(kept);
}

/**
 * The serializer a copy is written with.
 *
 * `stampCopy` is handed the whole copy once it stands, and only once: `DOMSerializer` draws a
 * node's children by calling `serializeFragment` again with the element it drew as the target, so
 * a call carrying one is inside the copy rather than around it.
 */
export function clipboardSerializer(
  stampCopy?: (copy: HTMLElement | DocumentFragment) => void
): DOMSerializer {
  const nodes: DOMSerializer["nodes"] = Object.fromEntries(
    Object.entries(CLIPBOARD_NODES).flatMap(([name, spec]) => {
      const draw = spec.toClipboardDOM;
      // A node that draws nothing is taken out of the fragment before the serializer sees it, so
      // the empty span below stands for a case `written` has already ruled out
      return draw === null
        ? []
        : [[name, (node: PMNode) => draw(node) ?? ["span"]]];
    })
  );
  const marks: DOMSerializer["marks"] = Object.fromEntries(
    Object.entries(CLIPBOARD_MARKS).flatMap(([name, spec]) =>
      // A mark with no drawing of its own is left out, and `DOMSerializer` writes the content it
      // covered without a wrapper
      spec.toClipboardDOM === null ? [] : [[name, spec.toClipboardDOM]]
    )
  );
  return new (class extends DOMSerializer {
    serializeFragment(
      fragment: Fragment,
      options?: { document?: Document },
      target?: HTMLElement | DocumentFragment
    ): HTMLElement | DocumentFragment {
      const dom = super.serializeFragment(written(fragment), options, target);
      if (target === undefined) stampCopy?.(dom);
      return dom;
    }
  })(nodes, marks);
}

function textPieces(fragment: Fragment): string[] {
  const pieces: string[] = [];
  fragment.forEach((child) => {
    const piece = clipboardTextOf(child);
    if (piece !== null) pieces.push(piece);
  });
  return pieces;
}

function clipboardTextOf(node: PMNode): string | null {
  const spec = nodeSpec(node);
  if (spec === undefined || spec.toClipboardText === null) return null;
  return spec.toClipboardText(node, textPieces(node.content));
}

/**
 * The plain text a copy leaves beside the HTML.
 *
 * ProseMirror's own answer is the text content with a line between blocks, which loses a tab, a
 * page break, and the difference between the next cell and the next row. Somewhere those are the
 * whole of what was copied, a table pasted into a spreadsheet above all.
 */
export function clipboardText(slice: Slice): string {
  const pieces = textPieces(slice.content);
  return pieces.join(slice.content.firstChild?.isInline === true ? "" : "\n");
}
