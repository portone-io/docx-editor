/**
 * The React editor surface: the editor component, the download helper, and the
 * document structure and format values those two hand out.
 *
 * The engine underneath is the `./core` entry, which knows nothing about React
 * or the DOM. The ProseMirror handles re-exported here are the ones from the
 * copy the editor itself runs on, which is the consumer's own copy: the packages
 * are peer dependencies.
 * `docs/architecture.md` has the line between the two entries and why the one
 * copy matters.
 */

export type { Command } from "prosemirror-state";
export {
  EditorState,
  Plugin,
  PluginKey,
  TextSelection,
} from "prosemirror-state";
export { EditorView } from "prosemirror-view";
export type {
  DocxEditorHandle,
  DocxEditorMode,
  DocxEditorProps,
} from "./DocxEditor";
export { DocxEditor } from "./DocxEditor";
export type { DocxBytes, DocxSource } from "./docx/importDocx";
export type {
  DownloadDocxOptions,
  DownloadDocxResult,
} from "./download";
export { downloadDocx } from "./download";
export type { CommentAuthor } from "./editor/commands/commentCommands";
export type {
  CellFormat,
  CellVerticalAlign,
  DocumentDefaults,
  HighlightName,
  LineSpacing,
  NumberingRef,
  ParagraphAlign,
  ParagraphFormat,
  RowFormat,
  RowHeight,
  RunFormat,
  TableFormat,
  TableWidth,
  TableWidthType,
  UnderlineKind,
  VerticalAlign,
} from "./model/format";
export type {
  DocxExportErrorCode,
  DocxImportErrorCode,
} from "./ooxml/errors";
export { DocxExportError, DocxImportError } from "./ooxml/errors";
export { docxSchema } from "./schema";
/** Whose comments the panel offers to edit, which `DocxEditorMode` takes */
export type { EditableComments } from "./schema/protection";
/**
 * The lists the built-in pickers offer. They are exported so that a toolbar of your own can
 * offer the same ones without writing them out again; the pickers themselves read these very
 * lists, so there is only ever one of each.
 * A list of your own goes back the other way through `presets` on `DocxEditor`, which the
 * built-in pickers then offer in place of these.
 */
export type { ColorRow } from "./styles/colors";
export { DEFAULT_COLORS } from "./styles/colors";
export { DEFAULT_FONT_SIZES } from "./styles/fontSizes";
export type { FontFallbackGroup, FontFallbacks } from "./styles/fontStack";
export { DEFAULT_FONT_FALLBACKS } from "./styles/fontStack";
export { DEFAULT_FONTS } from "./styles/fonts";
export type { LineSpacingOption } from "./styles/lineSpacings";
export { DEFAULT_LINE_SPACINGS } from "./styles/lineSpacings";
export type { CellBorderOption } from "./table/cellBorders";
export { DEFAULT_CELL_BORDERS } from "./table/cellBorders";
export type { DocxEditorPresets } from "./ui/presets";
export type { DocxEditorZoom } from "./ui/zoom";
export { DEFAULT_ZOOM_LEVELS } from "./ui/zoom";
