/**
 * ProseMirror schema joining preserved DOCX fragments with validated display values. Original XML
 * remains on nodes and marks so untouched content can round-trip verbatim.
 *
 * Every block sits in one of two groups beside `block`. A `modelled` block is one the editor
 * takes apart and writes back out itself, so the writer decides what it says; a `preserved` one
 * goes back out as the XML it arrived as. Which of the two a block is decides how it is written
 * (`docx/serializeBlock`) and how a submitted file is compared against the original
 * (`docx/storyProjection`), so a new block node has to name one of them.
 *
 * `./attrRoles` declares each attr's provenance and comparison role; `attrClasses.test.ts`
 * checks coverage. The plugin guide defines the supported public surface.
 */

import { Schema } from "prosemirror-model";
import {
  spanCount,
  toBandSizes,
  toCellFormat,
  toCellMargins,
  toColWidth,
  toGridCols,
  toInsideBorders,
  toParagraphFormat,
  toRowFormat,
  toRunFormat,
  toTableFormat,
  toTableStyleConditions,
  toTableWidth,
} from "../model/format";
import {
  ANY_ELEMENT,
  ATTRIBUTES,
  acceptRawXml,
  ELEMENT,
  type RawXmlShape,
} from "../ooxml/fragment";
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

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return typeof value === "object" && value !== null;
}

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

/**
 * The key of the block this came from (`docx/session`), read back as it stands.
 *
 * The nodes whose selector is the attribute itself write an absent key as the empty string, so
 * that reads back as no key rather than as a block of the open document.
 */
function srcIdOf(dom: HTMLElement): string | null {
  const raw = dom.getAttribute("data-src");
  return raw === null || raw === "" ? null : raw;
}

/**
 * Whether every reply a comment carries holds the XML the comment parts take it back as.
 * A reply body goes back out into `word/comments.xml` (`docx/comments/writing`) rather than into
 * the story, so one smuggling a sibling would write it there.
 */
function repliesHoldTheirXml(value: unknown): boolean {
  if (!Array.isArray(value)) return true;
  const replies: readonly unknown[] = value;
  return replies.every((reply) => {
    if (!isRecord(reply)) return true;
    return (
      acceptRawXml(ELEMENT("comment"), text(reply.commentXml) ?? null) !==
        false &&
      acceptRawXml(ANY_ELEMENT, text(reply.extensionXml) ?? null) !== false
    );
  });
}

/**
 * A content control's opening tag, with everything it wrapped cut away.
 * `docx/serializeParagraph` and `docx/serializeTable` put back exactly the closing text named here,
 * and `docx/sdt` cuts the tag at the `w:sdtContent` those two write, so the properties are all that
 * may still hang off it.
 */
const SDT_PREFIX: RawXmlShape = {
  kind: "openTag",
  name: "sdt",
  closedBy: "<w:sdtContent/></w:sdt>",
  head: ["sdtPr", "sdtEndPr"],
};

/** A hyperlink's opening tag, closed the same way. `docx/hyperlink` cuts it at the tag itself */
const LINK_PREFIX: RawXmlShape = {
  kind: "openTag",
  name: "hyperlink",
  closedBy: "</w:hyperlink>",
  head: [],
};

/**
 * The raw XML an attr carries, checked against the shape that attr goes back out as.
 *
 * `false` for a fragment that does not hold it, which every rule reading one answers `false` to in
 * turn: the node or mark is then not read at all and its content settles one level plainer, rather
 * than a string the writer would splice into the file arriving in the document.
 */
function rawXml(
  dom: HTMLElement,
  attribute: string,
  shape: RawXmlShape
): string | null | false {
  return acceptRawXml(shape, dom.getAttribute(attribute));
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
    doc: {
      content: "block+",
      attrs: {
        /**
         * The definitions of the lists started while editing, which the export writes into
         * numbering.xml. Read back through `numbering/listRegistry`.
         */
        newLists: { default: null },
      },
    },
    paragraph: {
      group: "block modelled",
      content: "inline*",
      // Leading and trailing whitespace carries meaning in docx text, so it is not stripped when the DOM is read back
      whitespace: "pre",
      attrs: {
        srcId: { default: null },
        /** The attribute string of the `<w:p>` opening tag. null when there is none */
        pAttrs: { default: null },
        /** The whole `<w:pPr>...</w:pPr>` XML. null when there is none */
        pPr: { default: null },
        /** Derived paragraph formatting; see `./attrRoles`. */
        format: { default: null },
        /**
         * Derived character formatting drawn on the paragraph so unmarked text inherits it.
         * Its comparison role and provenance are declared in `./attrRoles`.
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
            "data-src": text(node.attrs.srcId),
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
          getAttrs: (dom) => {
            const pAttrs = rawXml(dom, "data-pattrs", ATTRIBUTES);
            const pPr = rawXml(dom, "data-ppr", ELEMENT("pPr"));
            if (pAttrs === false || pPr === false) return false;
            return {
              srcId: srcIdOf(dom),
              pAttrs,
              pPr,
              format: toParagraphFormat(
                parseJson(dom.getAttribute("data-fmt"))
              ),
              styleRun: toRunFormat(
                parseJson(dom.getAttribute("data-style-run"))
              ),
            };
          },
        },
      ],
    },
    /**
     * A table carries the original `<w:tblPr>` and `<w:tblGrid>` untouched, and when an
     * edited table is written back the new XML is built from those two.
     */
    table: {
      group: "block modelled",
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
        /**
         * The whole `<w:tblGridChange>...</w:tblGridChange>` XML, the record of the grid this
         * table had before it was last revised. Carried as it arrived because the grid around it
         * is rebuilt from `gridCols`, and CT_TblGrid takes it after the columns however wide
         * those turn out to be (ECMA-376 Part 1 17.4.48)
         */
        gridChange: { default: null },
        format: { default: null },
        /** The lines between cells the table style laid down, so an edit can derive them again */
        styleInside: { default: null },
        /** The cell margins the table style laid down, carried along for the same reason */
        styleCellMargins: { default: null },
        /**
         * What the table style dresses each part of a table with (`w:tblStylePr`), by the part it
         * covers. Which of them a cell takes is worked out from where it sits and the table's own
         * `w:tblLook`, so an edit that moves the grid derives them again
         */
        styleConditions: { default: null },
        /** How many rows and columns one band of the table style is made of */
        styleBands: { default: null },
      },
      toDOM(node) {
        const format = toTableFormat(node.attrs.format);
        const width = toTableWidth(node.attrs.tblW);
        const gridCols = toGridCols(node.attrs.gridCols);
        const attrs = {
          class: editorClassNames.table,
          style: tableStyle(format, width),
          "data-src": text(node.attrs.srcId),
          "data-tblattrs": text(node.attrs.tblAttrs),
          "data-tblpr": text(node.attrs.tblPr),
          "data-tblw": formatJson(width),
          "data-cols": numberListText(gridCols),
          "data-gridchange": text(node.attrs.gridChange),
          "data-fmt": formatJson(format),
          "data-style-inside": formatJson(
            toInsideBorders(node.attrs.styleInside)
          ),
          "data-style-margins": formatJson(
            toCellMargins(node.attrs.styleCellMargins)
          ),
          "data-style-conditions": formatJson(
            toTableStyleConditions(node.attrs.styleConditions)
          ),
          "data-style-bands": formatJson(toBandSizes(node.attrs.styleBands)),
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
          getAttrs: (dom) => {
            const tblAttrs = rawXml(dom, "data-tblattrs", ATTRIBUTES);
            const tblPr = rawXml(dom, "data-tblpr", ELEMENT("tblPr"));
            const gridChange = rawXml(
              dom,
              "data-gridchange",
              ELEMENT("tblGridChange")
            );
            if (tblAttrs === false || tblPr === false || gridChange === false) {
              return false;
            }
            return {
              srcId: srcIdOf(dom),
              tblAttrs,
              tblPr,
              tblW: toTableWidth(parseJson(dom.getAttribute("data-tblw"))),
              gridCols: parseNumberList(dom.getAttribute("data-cols")),
              gridChange,
              format: toTableFormat(parseJson(dom.getAttribute("data-fmt"))),
              styleInside: toInsideBorders(
                parseJson(dom.getAttribute("data-style-inside"))
              ),
              styleCellMargins: toCellMargins(
                parseJson(dom.getAttribute("data-style-margins"))
              ),
              styleConditions: toTableStyleConditions(
                parseJson(dom.getAttribute("data-style-conditions"))
              ),
              styleBands: toBandSizes(
                parseJson(dom.getAttribute("data-style-bands"))
              ),
            };
          },
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
          getAttrs: (dom) => {
            const trAttrs = rawXml(dom, "data-trattrs", ATTRIBUTES);
            const tblPrEx = rawXml(dom, "data-tblprex", ELEMENT("tblPrEx"));
            const trPr = rawXml(dom, "data-trpr", ELEMENT("trPr"));
            if (trAttrs === false || tblPrEx === false || trPr === false) {
              return false;
            }
            return {
              trAttrs,
              tblPrEx,
              trPr,
              format: toRowFormat(parseJson(dom.getAttribute("data-fmt"))),
            };
          },
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
          getAttrs: (dom) => {
            const tcAttrs = rawXml(dom, "data-tcattrs", ATTRIBUTES);
            const tcPr = rawXml(dom, "data-tcpr", ELEMENT("tcPr"));
            const sdtPrefix = rawXml(dom, "data-sdt-prefix", SDT_PREFIX);
            if (tcAttrs === false || tcPr === false || sdtPrefix === false) {
              return false;
            }
            return {
              colspan: parseInt10(dom.getAttribute("colspan"), 1),
              rowspan: parseInt10(dom.getAttribute("rowspan"), 1),
              colwidth: toColWidth(
                parseNumberList(dom.getAttribute("data-colwidth"))
              ),
              tcAttrs,
              tcPr,
              tcW: toTableWidth(parseJson(dom.getAttribute("data-tcw"))),
              format: toCellFormat(parseJson(dom.getAttribute("data-fmt"))),
              sdtPrefix,
              sdtContentsLocked:
                dom.getAttribute("data-sdt-contents-locked") === "1",
              sdtDeletionLocked:
                dom.getAttribute("data-sdt-deletion-locked") === "1",
            };
          },
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
      group: "block preserved",
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
          getAttrs: (dom) => {
            const xml = rawXml(dom, "data-xml", ANY_ELEMENT);
            if (xml === false) return false;
            return { xml, name: dom.getAttribute("data-name") };
          },
        },
      ],
    },
    /** The node that carries a non-paragraph body block (a table we could not model, sectPr, an unknown element) exactly as it came */
    docxRaw: {
      group: "block preserved",
      atom: true,
      selectable: false,
      attrs: {
        srcId: { default: null },
        /** The original element name (for example `w:tbl`) */
        name: { default: null },
      },
      toDOM(node) {
        const { className, label } = rawBlockDom(text(node.attrs.name));
        // This is the mark that identifies the node when the DOM is read back, so it is attached even without a block key
        return [
          "div",
          {
            class: className,
            "data-src": text(node.attrs.srcId) ?? "",
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
      group: "block preserved",
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
            "data-src": text(node.attrs.srcId) ?? "",
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
          getAttrs: (dom) => {
            const brAttrs = rawXml(dom, "data-battrs", ATTRIBUTES);
            return brAttrs === false ? false : { brAttrs };
          },
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
          getAttrs: (dom) => {
            const xml = rawXml(dom, "data-xml", ELEMENT("drawing"));
            if (xml === false) return false;
            return {
              src: toImageSrc(dom.getAttribute("src")),
              extent: toImageExtent(parseJson(dom.getAttribute("data-extent"))),
              alt: dom.getAttribute("alt") || null,
              xml,
            };
          },
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
          getAttrs: (dom) => {
            const xml = rawXml(dom, "data-xml", ELEMENT("commentRangeStart"));
            if (xml === false) return false;
            return { id: dom.getAttribute("data-comment-id"), xml };
          },
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
          getAttrs: (dom) => {
            const xml = rawXml(dom, "data-xml", ELEMENT("commentRangeEnd"));
            if (xml === false) return false;
            return { id: dom.getAttribute("data-comment-id"), xml };
          },
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
        // The identity behind the display name, read from and written to the people part. Null for
        // a comment nobody's identity is recorded for, which every comment made in Word is here
        authorId: { default: null },
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
            "data-comment-author-id": text(node.attrs.authorId),
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
          getAttrs: (dom) => {
            const referenceXml = rawXml(
              dom,
              "data-reference-xml",
              ELEMENT("commentReference")
            );
            const commentXml = rawXml(
              dom,
              "data-comment-xml",
              ELEMENT("comment")
            );
            // The extended properties are a `w15:commentEx`, and `w15` is a namespace nothing
            // this far down knows, so the element is held to its shape alone
            const extensionXml = rawXml(
              dom,
              "data-comment-extension-xml",
              ANY_ELEMENT
            );
            const replies =
              parseJson(dom.getAttribute("data-comment-replies")) ?? [];
            if (
              referenceXml === false ||
              commentXml === false ||
              extensionXml === false ||
              !repliesHoldTheirXml(replies)
            ) {
              return false;
            }
            return {
              id: dom.getAttribute("data-comment-id"),
              referenceXml,
              author: dom.getAttribute("data-comment-author"),
              authorId: dom.getAttribute("data-comment-author-id"),
              initials: dom.getAttribute("data-comment-initials"),
              date: dom.getAttribute("data-comment-date"),
              text: dom.getAttribute("data-comment-text") ?? "",
              commentXml,
              imported: dom.getAttribute("data-comment-imported") === "1",
              paraId: dom.getAttribute("data-comment-para-id"),
              resolved: dom.getAttribute("data-comment-resolved") === "1",
              extensionXml,
              threadImported:
                dom.getAttribute("data-comment-thread-imported") === "1",
              replies,
            };
          },
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
          getAttrs: (dom) => {
            const referenceXml = rawXml(
              dom,
              "data-reference-xml",
              ELEMENT("footnoteReference", "endnoteReference")
            );
            if (referenceXml === false) return false;
            return {
              kind:
                dom.getAttribute("data-note-kind") === "endnote"
                  ? "endnote"
                  : "footnote",
              id: dom.getAttribute("data-note-id"),
              label: dom.getAttribute("data-note-label") ?? "?",
              text: dom.getAttribute("data-note-text") ?? "",
              customMarkFollows:
                dom.getAttribute("data-custom-mark-follows") === "1",
              referenceXml,
            };
          },
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
          getAttrs: (dom) => {
            const xml = rawXml(dom, "data-xml", ANY_ELEMENT);
            return xml === false ? false : { xml };
          },
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
            const prefix = rawXml(dom, "data-sdt-prefix", SDT_PREFIX);
            // With no opening tag to put back there is no control left to write out
            if (prefix === null || prefix === false) return false;
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
            const prefix = rawXml(dom, "data-link-prefix", LINK_PREFIX);
            const href = dom.getAttribute("data-href");
            if (prefix === false) return false;
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
        /** Derived run formatting; see `./attrRoles`. */
        format: { default: null },
      },
      toDOM(mark) {
        return runMarkSpec(mark.attrs, DEFAULT_FONT_FALLBACKS);
      },
      parseDOM: [
        {
          tag: `span.${editorClassNames.run}`,
          getAttrs: (dom) => {
            const rAttrs = rawXml(dom, "data-rattrs", ATTRIBUTES);
            const rPr = rawXml(dom, "data-rpr", ELEMENT("rPr"));
            if (rAttrs === false || rPr === false) return false;
            return {
              rAttrs,
              rPr,
              format: toRunFormat(parseJson(dom.getAttribute("data-fmt"))),
            };
          },
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
          getAttrs: (dom) => {
            const tabAttrs = rawXml(dom, "data-tattrs", ATTRIBUTES);
            return tabAttrs === false ? false : { tabAttrs };
          },
        },
      ],
    },
  },
});
