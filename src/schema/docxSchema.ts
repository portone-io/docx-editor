/**
 * ProseMirror schema joining preserved DOCX fragments with validated display values. Original XML
 * remains on nodes and marks so untouched content can round-trip verbatim.
 */

import { Schema } from "prosemirror-model";
import {
  spanCount,
  toCellFormat,
  toCellMargins,
  toColWidth,
  toGridCols,
  toInsideBorders,
  toParagraphFormat,
  toRowFormat,
  toRunFormat,
  toTableFormat,
  toTableWidth,
} from "../model/format";
import { toImageExtent, toImageSrc } from "../ooxml/image";
import { editorAttributes, editorClassNames } from "../styles/classNames";
import { DEFAULT_FONT_FALLBACKS } from "../styles/fontStack";
import {
  cellStyle,
  columnWidthPx,
  paragraphStyle,
  rowStyle,
  tableStyle,
} from "../styles/inlineStyle";
import { imageNodeSpec, runMarkSpec } from "./rendering";

export { imageNodeSpec, runMarkSpec } from "./rendering";

function text(value: unknown): string | undefined {
  return typeof value === "string" ? value : undefined;
}

function numberText(value: unknown): string | undefined {
  return typeof value === "number" ? `${value}` : undefined;
}

function parseInt10(value: string | null, fallback: number): number {
  if (value === null) return fallback;
  const parsed = Number.parseInt(value, 10);
  return Number.isNaN(parsed) ? fallback : parsed;
}

/** Reads back column widths recorded in the `1000,1500` form */
function parseNumberList(value: string | null): number[] {
  if (!value) return [];
  return value
    .split(",")
    .map((entry) => Number.parseFloat(entry))
    .filter((entry) => Number.isFinite(entry));
}

function numberListText(values: number[]): string | undefined {
  return values.length > 0 ? values.join(",") : undefined;
}

function parseJson(value: string | null): unknown {
  if (!value) return null;
  try {
    const parsed: unknown = JSON.parse(value);
    return parsed;
  } catch {
    return null;
  }
}

function formatJson(format: object | null): string | undefined {
  return format ? JSON.stringify(format) : undefined;
}

/** Whether this is a page-splitting break (`<w:br w:type="page"/>`). Read straight from the original attribute text */
export function isPageBreak(brAttrs: unknown): boolean {
  return (
    typeof brAttrs === "string" && /(^|\s)(\w+:)?type="page"/.test(brAttrs)
  );
}

function srcIdOf(dom: HTMLElement): number | null {
  const raw = dom.getAttribute("data-src");
  if (raw === null) return null;
  const parsed = Number.parseInt(raw, 10);
  return Number.isNaN(parsed) ? null : parsed;
}

const TABLE_PLACEHOLDER = "Table (unsupported layout, original is preserved)";
const UNKNOWN_PLACEHOLDER =
  "Unsupported content (not editable, original is preserved)";

const TABLE_NAMES = ["w:tbl", "tbl"];

function rawBlockDom(name: string | undefined): {
  className: string;
  label: string;
} {
  if (name !== undefined && TABLE_NAMES.includes(name)) {
    return {
      className: editorClassNames.tablePlaceholder,
      label: TABLE_PLACEHOLDER,
    };
  }
  return { className: editorClassNames.rawBlock, label: UNKNOWN_PLACEHOLDER };
}

/**
 * Looking at `data-src` alone would turn a stray `div` dragged in or pasted from
 * outside into a preserved block pointing at an original fragment that does not
 * exist, so the class we attach is checked along with it.
 */
const DOCX_RAW_SELECTOR = [
  editorClassNames.rawBlock,
  editorClassNames.tablePlaceholder,
]
  .map((className) => `div[data-src].${className}`)
  .join(", ");

export const docxSchema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      group: "block",
      content: "inline*",
      // Leading and trailing whitespace carries meaning in docx text, so it is not stripped when the DOM is read back
      whitespace: "pre",
      attrs: {
        srcId: { default: null },
        /** The attribute string of the `<w:p>` opening tag. null when there is none */
        pAttrs: { default: null },
        /** The whole `<w:pPr>...</w:pPr>` XML. null when there is none */
        pPr: { default: null },
        /** The display values derived from reading pPr */
        format: { default: null },
        /**
         * The character formatting the style this paragraph wears lays down, drawn as the
         * paragraph's own CSS so that text carrying no run of its own inherits it.
         * It is derived from the style table the same way `format` is, and like `format` it never
         * goes back into the document.
         */
        styleRun: { default: null },
      },
      toDOM(node) {
        const format = toParagraphFormat(node.attrs.format);
        const styleRun = toRunFormat(node.attrs.styleRun);
        return [
          "p",
          {
            class: editorClassNames.paragraph,
            style: paragraphStyle(format, styleRun),
            "data-src":
              node.attrs.srcId === null ? undefined : `${node.attrs.srcId}`,
            "data-pattrs": text(node.attrs.pAttrs),
            "data-ppr": text(node.attrs.pPr),
            "data-fmt": formatJson(format),
            "data-style-run": formatJson(styleRun),
            [editorAttributes.pageBreakBefore]: format?.pageBreakBefore
              ? ""
              : undefined,
          },
          0,
        ];
      },
      parseDOM: [
        {
          tag: `p.${editorClassNames.paragraph}`,
          getAttrs: (dom) => ({
            srcId: srcIdOf(dom),
            pAttrs: dom.getAttribute("data-pattrs"),
            pPr: dom.getAttribute("data-ppr"),
            format: toParagraphFormat(parseJson(dom.getAttribute("data-fmt"))),
            styleRun: toRunFormat(
              parseJson(dom.getAttribute("data-style-run"))
            ),
          }),
        },
      ],
    },
    /**
     * A table carries the original `<w:tblPr>` and `<w:tblGrid>` untouched, and when an
     * edited table is written back the new XML is built from those two.
     */
    table: {
      group: "block",
      content: "tableRow+",
      tableRole: "table",
      isolating: true,
      attrs: {
        srcId: { default: null },
        /** The attribute string of the `<w:tbl>` opening tag */
        tblAttrs: { default: null },
        /** The whole `<w:tblPr>...</w:tblPr>` XML */
        tblPr: { default: null },
        /** The table width read out of tblPr */
        tblW: { default: null },
        /** The `w:gridCol` widths (dxa) in order */
        gridCols: { default: [] },
        format: { default: null },
        /** The lines between cells the table style laid down, so an edit can derive them again */
        styleInside: { default: null },
        /** The cell margins the table style laid down, carried along for the same reason */
        styleCellMargins: { default: null },
      },
      toDOM(node) {
        const format = toTableFormat(node.attrs.format);
        const width = toTableWidth(node.attrs.tblW);
        const gridCols = toGridCols(node.attrs.gridCols);
        const attrs = {
          class: editorClassNames.table,
          style: tableStyle(format, width),
          "data-src":
            node.attrs.srcId === null ? undefined : `${node.attrs.srcId}`,
          "data-tblattrs": text(node.attrs.tblAttrs),
          "data-tblpr": text(node.attrs.tblPr),
          "data-tblw": formatJson(width),
          "data-cols": numberListText(gridCols),
          "data-fmt": formatJson(format),
          "data-style-inside": formatJson(
            toInsideBorders(node.attrs.styleInside)
          ),
          "data-style-margins": formatJson(
            toCellMargins(node.attrs.styleCellMargins)
          ),
        };
        const body = ["tbody", 0];
        if (gridCols.length === 0) return ["table", attrs, body];
        const cols = gridCols.map((dxa) => [
          "col",
          { style: `width:${columnWidthPx(dxa)}px` },
        ]);
        return ["table", attrs, ["colgroup", ...cols], body];
      },
      parseDOM: [
        // `data-cols` carries the column widths, so the colgroup we drew is not read back
        { tag: "colgroup", ignore: true },
        {
          tag: `table.${editorClassNames.table}`,
          getAttrs: (dom) => ({
            srcId: srcIdOf(dom),
            tblAttrs: dom.getAttribute("data-tblattrs"),
            tblPr: dom.getAttribute("data-tblpr"),
            tblW: toTableWidth(parseJson(dom.getAttribute("data-tblw"))),
            gridCols: parseNumberList(dom.getAttribute("data-cols")),
            format: toTableFormat(parseJson(dom.getAttribute("data-fmt"))),
            styleInside: toInsideBorders(
              parseJson(dom.getAttribute("data-style-inside"))
            ),
            styleCellMargins: toCellMargins(
              parseJson(dom.getAttribute("data-style-margins"))
            ),
          }),
        },
      ],
    },
    tableRow: {
      content: "tableCell+",
      tableRole: "row",
      attrs: {
        /** The attribute string of the `<w:tr>` opening tag */
        trAttrs: { default: null },
        /** The whole `<w:tblPrEx>...</w:tblPrEx>` XML, carried along without being read */
        tblPrEx: { default: null },
        /** The whole `<w:trPr>...</w:trPr>` XML */
        trPr: { default: null },
        format: { default: null },
      },
      toDOM(node) {
        const format = toRowFormat(node.attrs.format);
        return [
          "tr",
          {
            class: editorClassNames.tableRow,
            style: rowStyle(format),
            "data-trattrs": text(node.attrs.trAttrs),
            "data-tblprex": text(node.attrs.tblPrEx),
            "data-trpr": text(node.attrs.trPr),
            "data-fmt": formatJson(format),
          },
          0,
        ];
      },
      parseDOM: [
        {
          tag: "tr",
          getAttrs: (dom) => ({
            trAttrs: dom.getAttribute("data-trattrs"),
            tblPrEx: dom.getAttribute("data-tblprex"),
            trPr: dom.getAttribute("data-trpr"),
            format: toRowFormat(parseJson(dom.getAttribute("data-fmt"))),
          }),
        },
      ],
    },
    /**
     * A vertically merged cell exists only as the one cell that starts the merge, and
     * `rowspan` counts how many rows it covers.
     * The empty cells on the continuing rows are rebuilt on export.
     */
    tableCell: {
      content: "block+",
      tableRole: "cell",
      isolating: true,
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        colwidth: { default: null },
        /** The attribute string of the `<w:tc>` opening tag */
        tcAttrs: { default: null },
        /** The whole `<w:tcPr>...</w:tcPr>` XML. It still holds the original vMerge and gridSpan */
        tcPr: { default: null },
        /** The cell width read out of tcPr */
        tcW: { default: null },
        format: { default: null },
        /** The opening XML of the `<w:sdt>` content control this cell sat inside */
        sdtPrefix: { default: null },
        /**
         * The two clauses of that control's lock, which mean the same as `contentsLocked` and
         * `deletionLocked` on the sdt mark. Both only ever stand alongside an `sdtPrefix`, since
         * the lock lives inside that very XML.
         */
        sdtContentsLocked: { default: false },
        sdtDeletionLocked: { default: false },
      },
      toDOM(node) {
        const format = toCellFormat(node.attrs.format);
        const colspan = spanCount(node.attrs.colspan);
        const rowspan = spanCount(node.attrs.rowspan);
        const colwidth = toColWidth(node.attrs.colwidth);
        const locked = node.attrs.sdtContentsLocked === true;
        return [
          "td",
          {
            class: locked
              ? `${editorClassNames.tableCell} ${editorClassNames.cellLocked}`
              : editorClassNames.tableCell,
            style: cellStyle(format),
            colspan: colspan > 1 ? `${colspan}` : undefined,
            rowspan: rowspan > 1 ? `${rowspan}` : undefined,
            "data-colwidth": colwidth ? numberListText(colwidth) : undefined,
            "data-tcattrs": text(node.attrs.tcAttrs),
            "data-tcpr": text(node.attrs.tcPr),
            "data-tcw": formatJson(toTableWidth(node.attrs.tcW)),
            "data-fmt": formatJson(format),
            "data-sdt-prefix": text(node.attrs.sdtPrefix),
            "data-sdt-contents-locked": locked ? "1" : undefined,
            "data-sdt-deletion-locked":
              node.attrs.sdtDeletionLocked === true ? "1" : undefined,
          },
          0,
        ];
      },
      parseDOM: [
        {
          tag: "td",
          getAttrs: (dom) => ({
            colspan: parseInt10(dom.getAttribute("colspan"), 1),
            rowspan: parseInt10(dom.getAttribute("rowspan"), 1),
            colwidth: toColWidth(
              parseNumberList(dom.getAttribute("data-colwidth"))
            ),
            tcAttrs: dom.getAttribute("data-tcattrs"),
            tcPr: dom.getAttribute("data-tcpr"),
            tcW: toTableWidth(parseJson(dom.getAttribute("data-tcw"))),
            format: toCellFormat(parseJson(dom.getAttribute("data-fmt"))),
            sdtPrefix: dom.getAttribute("data-sdt-prefix"),
            sdtContentsLocked:
              dom.getAttribute("data-sdt-contents-locked") === "1",
            sdtDeletionLocked:
              dom.getAttribute("data-sdt-deletion-locked") === "1",
          }),
        },
      ],
    },
    /**
     * The node that carries, as original XML, a block inside a table cell that we do
     * not model (a nested table, for instance).
     * Unlike a body block it has no original fragment number, so it carries its own
     * XML around with it.
     */
    rawBlock: {
      group: "block",
      atom: true,
      selectable: false,
      attrs: {
        xml: { default: null },
        /** The original element name (for example `w:tbl`) */
        name: { default: null },
      },
      toDOM(node) {
        return [
          "div",
          {
            class: `${editorClassNames.rawBlock} ${editorClassNames.rawXmlBlock}`,
            "data-xml": text(node.attrs.xml),
            "data-name": text(node.attrs.name),
          },
          UNKNOWN_PLACEHOLDER,
        ];
      },
      parseDOM: [
        {
          tag: `div.${editorClassNames.rawXmlBlock}`,
          getAttrs: (dom) => ({
            xml: dom.getAttribute("data-xml"),
            name: dom.getAttribute("data-name"),
          }),
        },
      ],
    },
    /** The node that carries a non-paragraph body block (a table we could not model, sectPr, an unknown element) exactly as it came */
    docxRaw: {
      group: "block",
      atom: true,
      selectable: false,
      attrs: {
        srcId: { default: null },
        /** The original element name (for example `w:tbl`) */
        name: { default: null },
      },
      toDOM(node) {
        const { className, label } = rawBlockDom(text(node.attrs.name));
        // This is the mark that identifies the node when the DOM is read back, so it is attached even without an original number
        return [
          "div",
          {
            class: className,
            "data-src": node.attrs.srcId === null ? "" : `${node.attrs.srcId}`,
            "data-name": text(node.attrs.name),
          },
          label,
        ];
      },
      parseDOM: [
        {
          tag: DOCX_RAW_SELECTOR,
          getAttrs: (dom) => ({
            srcId: srcIdOf(dom),
            name: dom.getAttribute("data-name"),
          }),
        },
      ],
    },
    /** A bookmark range marker that occurs directly under w:body rather than inside a paragraph */
    bookmarkBlock: {
      group: "block",
      atom: true,
      isolating: true,
      selectable: false,
      attrs: {
        srcId: { default: null },
        name: { default: null },
      },
      toDOM(node) {
        return [
          "div",
          {
            class: editorClassNames.bookmarkBlock,
            "data-src": node.attrs.srcId === null ? "" : `${node.attrs.srcId}`,
            "data-name": text(node.attrs.name),
            hidden: "hidden",
          },
        ];
      },
      parseDOM: [
        {
          tag: `div.${editorClassNames.bookmarkBlock}`,
          getAttrs: (dom) => ({
            srcId: srcIdOf(dom),
            name: dom.getAttribute("data-name"),
          }),
        },
      ],
    },
    hardBreak: {
      group: "inline",
      inline: true,
      marks: "run sdt link",
      attrs: { brAttrs: { default: null } },
      toDOM(node) {
        return [
          "br",
          {
            "data-battrs": text(node.attrs.brAttrs),
            [editorAttributes.breakType]: isPageBreak(node.attrs.brAttrs)
              ? "page"
              : undefined,
          },
        ];
      },
      parseDOM: [
        {
          tag: "br",
          getAttrs: (dom) => ({ brAttrs: dom.getAttribute("data-battrs") }),
        },
      ],
    },
    /**
     * `src` carries the bytes themselves as a data URL (see `docx/media` for why), and
     * `extent` the size the drawing records, in EMU.
     * An imported image also holds its whole original `<w:drawing>` XML, so an image
     * nobody touched goes back out byte for byte and a resize rewrites nothing but the
     * two extents. An image inserted during editing has no such XML, and the export
     * builds a drawing for it from the bytes.
     */
    image: {
      group: "inline",
      inline: true,
      draggable: true,
      marks: "run sdt link",
      attrs: {
        /** The image bytes as a data URL */
        src: { default: null },
        /** The display size in EMU, as `{ cx, cy }` */
        extent: { default: null },
        /** The alternative text from `wp:docPr descr`. null when there is none */
        alt: { default: null },
        /** The whole original `<w:drawing>` XML. null for an image inserted during editing */
        xml: { default: null },
      },
      toDOM(node) {
        return imageNodeSpec(node.attrs);
      },
      parseDOM: [
        {
          tag: `img.${editorClassNames.image}`,
          getAttrs: (dom) => ({
            src: toImageSrc(dom.getAttribute("src")),
            extent: toImageExtent(parseJson(dom.getAttribute("data-extent"))),
            alt: dom.getAttribute("alt") || null,
            xml: dom.getAttribute("data-xml"),
          }),
        },
      ],
    },
    commentStart: {
      group: "inline",
      inline: true,
      atom: true,
      selectable: false,
      marks: "sdt link",
      attrs: {
        id: { default: null },
        xml: { default: null },
      },
      toDOM(node) {
        return [
          "span",
          {
            class: editorClassNames.commentMarker,
            "data-comment-marker": "start",
            "data-comment-id": text(node.attrs.id),
            "data-xml": text(node.attrs.xml),
          },
        ];
      },
      parseDOM: [
        {
          tag: `span.${editorClassNames.commentMarker}[data-comment-marker="start"]`,
          getAttrs: (dom) => ({
            id: dom.getAttribute("data-comment-id"),
            xml: dom.getAttribute("data-xml"),
          }),
        },
      ],
    },
    commentEnd: {
      group: "inline",
      inline: true,
      atom: true,
      selectable: false,
      marks: "sdt link",
      attrs: {
        id: { default: null },
        xml: { default: null },
      },
      toDOM(node) {
        return [
          "span",
          {
            class: editorClassNames.commentMarker,
            "data-comment-marker": "end",
            "data-comment-id": text(node.attrs.id),
            "data-xml": text(node.attrs.xml),
          },
        ];
      },
      parseDOM: [
        {
          tag: `span.${editorClassNames.commentMarker}[data-comment-marker="end"]`,
          getAttrs: (dom) => ({
            id: dom.getAttribute("data-comment-id"),
            xml: dom.getAttribute("data-xml"),
          }),
        },
      ],
    },
    commentReference: {
      group: "inline",
      inline: true,
      atom: true,
      selectable: false,
      marks: "run sdt link",
      attrs: {
        id: { default: null },
        referenceXml: { default: null },
        author: { default: null },
        initials: { default: null },
        date: { default: null },
        text: { default: "" },
        commentXml: { default: null },
        imported: { default: false },
        paraId: { default: null },
        resolved: { default: false },
        extensionXml: { default: null },
        threadImported: { default: false },
        replies: { default: [] },
      },
      toDOM(node) {
        return [
          "span",
          {
            class: editorClassNames.commentMarker,
            "data-comment-marker": "reference",
            "data-comment-id": text(node.attrs.id),
            "data-reference-xml": text(node.attrs.referenceXml),
            "data-comment-author": text(node.attrs.author),
            "data-comment-initials": text(node.attrs.initials),
            "data-comment-date": text(node.attrs.date),
            "data-comment-text": text(node.attrs.text),
            "data-comment-xml": text(node.attrs.commentXml),
            "data-comment-imported":
              node.attrs.imported === true ? "1" : undefined,
            "data-comment-para-id": text(node.attrs.paraId),
            "data-comment-resolved":
              node.attrs.resolved === true ? "1" : undefined,
            "data-comment-extension-xml": text(node.attrs.extensionXml),
            "data-comment-thread-imported":
              node.attrs.threadImported === true ? "1" : undefined,
            "data-comment-replies": JSON.stringify(node.attrs.replies ?? []),
          },
        ];
      },
      parseDOM: [
        {
          tag: `span.${editorClassNames.commentMarker}[data-comment-marker="reference"]`,
          getAttrs: (dom) => ({
            id: dom.getAttribute("data-comment-id"),
            referenceXml: dom.getAttribute("data-reference-xml"),
            author: dom.getAttribute("data-comment-author"),
            initials: dom.getAttribute("data-comment-initials"),
            date: dom.getAttribute("data-comment-date"),
            text: dom.getAttribute("data-comment-text") ?? "",
            commentXml: dom.getAttribute("data-comment-xml"),
            imported: dom.getAttribute("data-comment-imported") === "1",
            paraId: dom.getAttribute("data-comment-para-id"),
            resolved: dom.getAttribute("data-comment-resolved") === "1",
            extensionXml: dom.getAttribute("data-comment-extension-xml"),
            threadImported:
              dom.getAttribute("data-comment-thread-imported") === "1",
            replies: parseJson(dom.getAttribute("data-comment-replies")) ?? [],
          }),
        },
      ],
    },
    noteReference: {
      group: "inline",
      inline: true,
      atom: true,
      selectable: false,
      marks: "run sdt link",
      attrs: {
        kind: { default: "footnote" },
        id: { default: null },
        label: { default: "?" },
        text: { default: "" },
        customMarkFollows: { default: false },
        referenceXml: { default: null },
      },
      toDOM(node) {
        const kind = node.attrs.kind === "endnote" ? "Endnote" : "Footnote";
        const label = text(node.attrs.label) ?? "?";
        const body = text(node.attrs.text) ?? "";
        return [
          "sup",
          {
            class: editorClassNames.noteReference,
            "data-note-kind": text(node.attrs.kind),
            "data-note-id": text(node.attrs.id),
            "data-note-label": label,
            "data-note-text": body,
            "data-custom-mark-follows":
              node.attrs.customMarkFollows === true ? "1" : undefined,
            "data-reference-xml": text(node.attrs.referenceXml),
            "aria-label": `${kind} ${label}`,
            title: body,
          },
          node.attrs.customMarkFollows === true ? "" : label,
        ];
      },
      parseDOM: [
        {
          tag: `sup.${editorClassNames.noteReference}`,
          getAttrs: (dom) => ({
            kind:
              dom.getAttribute("data-note-kind") === "endnote"
                ? "endnote"
                : "footnote",
            id: dom.getAttribute("data-note-id"),
            label: dom.getAttribute("data-note-label") ?? "?",
            text: dom.getAttribute("data-note-text") ?? "",
            customMarkFollows:
              dom.getAttribute("data-custom-mark-follows") === "1",
            referenceXml: dom.getAttribute("data-reference-xml"),
          }),
        },
      ],
    },
    /** The node that carries a non-run element inside a paragraph (a bookmark, for instance) exactly as it came */
    rawInline: {
      group: "inline",
      inline: true,
      atom: true,
      marks: "sdt link",
      attrs: { xml: { default: null } },
      toDOM(node) {
        return [
          "span",
          {
            class: editorClassNames.rawInline,
            "data-xml": text(node.attrs.xml),
          },
        ];
      },
      parseDOM: [
        {
          tag: `span.${editorClassNames.rawInline}`,
          getAttrs: (dom) => ({ xml: dom.getAttribute("data-xml") }),
        },
      ],
    },
    text: { group: "inline" },
  },
  marks: {
    /**
     * The content control (`w:sdt`) a stretch of inline content sits inside.
     *
     * It is declared ahead of the run mark on purpose: the mark declared first is drawn
     * outside on screen, so one control can hold the several runs it wrapped in the file.
     */
    sdt: {
      // A character typed against either edge of the control belongs outside it
      inclusive: false,
      attrs: {
        /** The opening XML of the `<w:sdt>`, the same string a wrapped cell carries */
        sdtPrefix: { default: null },
        /**
         * Which control this is, counted through the document as it was opened, and never
         * written back to the file. Two controls whose XML is identical would otherwise wear
         * the very same mark and their text would run into one. The first control gets 0, so
         * a control made during editing needs no number of its own.
         */
        sdtKey: { default: 0 },
        /**
         * The two clauses of the control's lock: whether its contents may not be edited, and
         * whether the control itself may not be deleted, not even whole. The `w:lock` inside
         * sdtPrefix is where both come from, and reading it once on import saves parsing that
         * string again on every draw.
         * The two are independent: a control may be un-editable yet removable, or editable yet
         * not removable.
         */
        contentsLocked: { default: false },
        deletionLocked: { default: false },
      },
      toDOM(mark) {
        const locked = mark.attrs.contentsLocked === true;
        return [
          "span",
          {
            class: locked
              ? `${editorClassNames.sdt} ${editorClassNames.sdtLocked}`
              : editorClassNames.sdt,
            "data-sdt-prefix": text(mark.attrs.sdtPrefix),
            "data-sdt-key": numberText(mark.attrs.sdtKey),
            "data-sdt-contents-locked": locked ? "1" : undefined,
            "data-sdt-deletion-locked":
              mark.attrs.deletionLocked === true ? "1" : undefined,
          },
          0,
        ];
      },
      parseDOM: [
        {
          tag: `span.${editorClassNames.sdt}`,
          getAttrs: (dom) => {
            const prefix = dom.getAttribute("data-sdt-prefix");
            // With no opening tag to put back there is no control left to write out
            if (prefix === null) return false;
            return {
              sdtPrefix: prefix,
              sdtKey: parseInt10(dom.getAttribute("data-sdt-key"), 0),
              contentsLocked:
                dom.getAttribute("data-sdt-contents-locked") === "1",
              deletionLocked:
                dom.getAttribute("data-sdt-deletion-locked") === "1",
            };
          },
        },
      ],
    },
    /**
     * The hyperlink (`w:hyperlink`) a stretch of inline content sits inside.
     *
     * It stands between the control and the run for the same reason the control stands outside
     * both: a link wraps whole runs in the file, and a link inside a content control has to come
     * back out inside it (`docx/serializeParagraph`).
     */
    link: {
      // A character typed against either edge of a link belongs outside it, and that is also what
      // decides whether a caret counts as standing in one (`editor/commands/linkCommands`)
      inclusive: false,
      attrs: {
        /**
         * The opening XML of the `<w:hyperlink>`, which carries everything about the link we never
         * read: `w:tooltip`, `w:history`, `w:docLocation`, `w:anchor`. null for a link made in the
         * editor, which the export writes an opening tag for.
         */
        linkPrefix: { default: null },
        /**
         * The address, read through the relationship the wrapper names. null for a link that names
         * a bookmark alone, or one whose relationship leads nowhere we can follow: the wrapper still
         * travels, and no address is offered for it.
         */
        href: { default: null },
        /**
         * Which link this is, counted through the document as it was opened, and never written back
         * to the file. Two links written exactly alike would otherwise wear the very same mark and
         * the text they cover would run into one wrapper. The first link gets 0, so a link made
         * during editing needs no number of its own.
         */
        linkKey: { default: 0 },
      },
      toDOM(mark) {
        return [
          "span",
          {
            class: editorClassNames.link,
            // The address is drawn as data rather than as an `href`, so that nothing on the page
            // navigates: a click in the text places the caret, and opening the address is offered
            // by the link panel instead
            "data-href": text(mark.attrs.href),
            "data-link-prefix": text(mark.attrs.linkPrefix),
            "data-link-key": numberText(mark.attrs.linkKey),
          },
          0,
        ];
      },
      parseDOM: [
        {
          tag: `span.${editorClassNames.link}`,
          getAttrs: (dom) => {
            const prefix = dom.getAttribute("data-link-prefix");
            const href = dom.getAttribute("data-href");
            // With neither an opening tag to put back nor an address to write one from, there is
            // no link left to write out
            if (prefix === null && href === null) return false;
            return {
              linkPrefix: prefix,
              href,
              linkKey: parseInt10(dom.getAttribute("data-link-key"), 0),
            };
          },
        },
      ],
    },
    /**
     * The run mark carries the original XML untouched, so even formatting we cannot
     * interpret survives editing unchanged.
     */
    run: {
      attrs: {
        rPr: { default: null },
        rAttrs: { default: null },
        /** The display values derived from reading rPr */
        format: { default: null },
      },
      toDOM(mark) {
        return runMarkSpec(mark.attrs, DEFAULT_FONT_FALLBACKS);
      },
      parseDOM: [
        {
          tag: `span.${editorClassNames.run}`,
          getAttrs: (dom) => ({
            rAttrs: dom.getAttribute("data-rattrs"),
            rPr: dom.getAttribute("data-rpr"),
            format: toRunFormat(parseJson(dom.getAttribute("data-fmt"))),
          }),
        },
      ],
    },
    /** A text tab whose mark retains the attributes of the source `w:tab`. */
    tab: {
      inclusive: false,
      attrs: { tabAttrs: { default: null } },
      toDOM(mark) {
        return [
          "span",
          {
            class: editorClassNames.tab,
            "data-tattrs": text(mark.attrs.tabAttrs),
          },
          0,
        ];
      },
      parseDOM: [
        {
          tag: `span.${editorClassNames.tab}`,
          getAttrs: (dom) => ({ tabAttrs: dom.getAttribute("data-tattrs") }),
        },
      ],
    },
  },
});
