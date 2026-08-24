/**
 * Downloads the document the editor holds as a docx file.
 *
 * The pieces it is built out of are exported, but not from the package's barrel:
 * `downloadDocx` is all a screen needs.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { DocxEditorHandle } from "./DocxEditor";

export const DOCX_MIME_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/**
 * How long the objectURL is kept alive so the browser can finish taking the file.
 *
 * Revoking it right after the click sometimes cancels the download on Firefox and
 * Safari.
 * All that leaks while the revoke is postponed is a single URL, so the window is
 * set generously.
 */
const REVOKE_DELAY_MS = 60_000;

export interface DownloadDocxOptions {
  /** The name to save under. The `.docx` suffix may be there or not */
  fileName: string;
}

export type DownloadDocxResult =
  | { status: "exported"; fileName: string; byteLength: number }
  | { status: "unavailable" }
  | { status: "empty" };

/** Settles the end of the name on exactly one `.docx` */
export function withDocxExtension(fileName: string): string {
  return `${fileName.replace(/\.docx$/i, "")}.docx`;
}

/**
 * Wraps the docx bytes into a single file blob.
 * A `Uint8Array` may be one slice of a larger buffer, so a copy is handed over.
 */
export function createDocxBlob(bytes: Uint8Array): Blob {
  return new Blob([new Uint8Array(bytes)], { type: DOCX_MIME_TYPE });
}

/**
 * Downloads a single file blob.
 *
 * The anchor is attached to the document and then taken off again because some
 * browsers only honor the click while it is attached.
 * Revoking the objectURL is deferred by `REVOKE_DELAY_MS`.
 */
export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.rel = "noopener";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), REVOKE_DELAY_MS);
}

/**
 * Whether there is anything worth exporting.
 *
 * Any text counts as content, and so does a block that shows on screen without text of
 * its own (a table, a placeholder).
 * A document of nothing but empty paragraphs would download as a file with nothing to
 * open, so it is taken to be empty.
 */
export function hasExportableContent(doc: PMNode): boolean {
  if (doc.textContent.length > 0) return true;
  let visible = false;
  doc.forEach((block) => {
    if (block.type.name !== "paragraph") visible = true;
  });
  return visible;
}

/**
 * Downloads the edited document as a docx file.
 *
 * `unavailable` when the editor is not mounted yet or the document could not be
 * opened, `empty` when there is nothing to export, and `exported` once a file has
 * gone out.
 * If writing the document back is blocked (a new list in a document without
 * numbering.xml, and the like), it is not passed over silently: `DocxExportError`
 * is rethrown as is.
 */
export function downloadDocx(
  editor: DocxEditorHandle | null | undefined,
  options: DownloadDocxOptions
): DownloadDocxResult {
  if (!editor) return { status: "unavailable" };
  if (!hasExportableContent(editor.view.state.doc)) return { status: "empty" };

  const bytes = editor.exportBytes();
  const fileName = withDocxExtension(options.fileName);
  downloadBlob(createDocxBlob(bytes), fileName);
  return { status: "exported", fileName, byteLength: bytes.length };
}
