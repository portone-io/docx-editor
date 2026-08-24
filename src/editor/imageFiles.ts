/**
 * Image files coming in from outside: from the picker the toolbar opens, from a paste, and
 * from a drop.
 *
 * All three take one path. The file is read as a data URL, the decoded image states its
 * natural pixel size, and that size - shrunk to the width of a line of body text when the
 * image is wider than the page - becomes the size the drawing records.
 * Only the kinds `docx/image` can write back are taken; a file of any other kind is left
 * where it was, for whoever else wants the event.
 */

import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { MAX_IMAGE_BYTES } from "../docx/media";
import { A4_BODY_WIDTH } from "../docx/pageGeometry";
import {
  emuToPx,
  IMAGE_MIMES,
  type ImageExtent,
  type ImageMime,
  imageBase64Of,
  isImageMime,
  pxToEmu,
  toImageSrc,
} from "../ooxml/image";
import { documentBodyWidthPx } from "./documentStyles";
import { type ImageToInsert, insertImage } from "./insertImage";
import { moveCaretToDrop } from "./plugins/dropCaret";

/** What a file picker is told to accept: the kinds we can both draw and write back */
export const IMAGE_FILE_ACCEPT = IMAGE_MIMES.join(",");

/** Anything holding a list of files: a clipboard, a drag, a file input */
interface FileCarrier {
  readonly files: ArrayLike<File> | null;
}

/** The image files a clipboard, a drag or an input is carrying, in the order they came */
export function imageFilesIn(carrier: FileCarrier | null | undefined): File[] {
  const files = carrier?.files;
  return files
    ? Array.from(files).filter((file) => isImageMime(file.type))
    : [];
}

/**
 * The size an image is first shown at.
 *
 * Its own pixel size, shrunk to the supplied width and height caps. A photo straight off a
 * camera would otherwise come in several times wider than the paper. An image that already
 * fits is never grown.
 * Null for a size we cannot show, so nothing goes in that would draw as an empty box.
 *
 * The width is the open document's own (`documentBodyWidthPx`). A caller with no document to
 * ask gets the A4 body width, which is the width every document used to be fitted to.
 */
export function fittedExtent(
  width: number,
  height: number,
  maxWidthPx: number = A4_BODY_WIDTH.px,
  maxHeightPx: number = Number.POSITIVE_INFINITY
): ImageExtent | null {
  if (
    !Number.isFinite(width) ||
    !Number.isFinite(height) ||
    !(width > 0) ||
    !(height > 0) ||
    !(maxWidthPx > 0) ||
    !(maxHeightPx > 0)
  ) {
    return null;
  }
  const scale = Math.min(1, maxWidthPx / width, maxHeightPx / height);
  const extent = {
    cx: pxToEmu(width * scale),
    cy: pxToEmu(height * scale),
  };
  return Number.isSafeInteger(extent.cx) &&
    Number.isSafeInteger(extent.cy) &&
    extent.cx > 0 &&
    extent.cy > 0
    ? extent
    : null;
}

/** The bytes of the file as a data URL. Null when it could not be read */
function readDataUrl(file: Blob): Promise<string | null> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.addEventListener("load", () =>
      resolve(typeof reader.result === "string" ? reader.result : null)
    );
    reader.addEventListener("error", () => resolve(null));
    reader.readAsDataURL(file);
  });
}

function imageByteLength(src: string): number {
  const base64 = imageBase64Of(src);
  const padding = base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0;
  return Math.floor((base64.length * 3) / 4) - padding;
}

function imageMimeFromBytes(binary: string): ImageMime | null {
  const byte = (index: number) => binary.charCodeAt(index);
  if (
    byte(0) === 0x89 &&
    binary.slice(1, 4) === "PNG" &&
    [0x0d, 0x0a, 0x1a, 0x0a].every((value, index) => byte(index + 4) === value)
  ) {
    return "image/png";
  }
  if (byte(0) === 0xff && byte(1) === 0xd8 && byte(2) === 0xff) {
    return "image/jpeg";
  }
  if (binary.startsWith("GIF87a") || binary.startsWith("GIF89a")) {
    return "image/gif";
  }
  return binary.startsWith("BM") ? "image/bmp" : null;
}

export function boundedImageSrc(source: unknown): string | null {
  const src = toImageSrc(source);
  if (src === null || imageByteLength(src) > MAX_IMAGE_BYTES) return null;
  try {
    const base64 = imageBase64Of(src);
    const mime = imageMimeFromBytes(atob(base64));
    return mime === null ? null : `data:${mime};base64,${base64}`;
  } catch {
    return null;
  }
}

export function imageSourceByteLength(source: string): number {
  return imageByteLength(source);
}

/** The decoded pixel size, or null when decoding fails or is canceled */
function naturalSize(
  src: string,
  signal?: AbortSignal
): Promise<{ width: number; height: number } | null> {
  return new Promise((resolve) => {
    const image = new Image();
    let settled = false;
    const finish = (value: { width: number; height: number } | null) => {
      if (settled) return;
      settled = true;
      signal?.removeEventListener("abort", onAbort);
      image.removeEventListener("load", onLoad);
      image.removeEventListener("error", onError);
      resolve(value);
    };
    const onLoad = () =>
      finish({ width: image.naturalWidth, height: image.naturalHeight });
    const onError = () => finish(null);
    const onAbort = () => {
      finish(null);
      image.src = "";
    };
    image.addEventListener("load", onLoad);
    image.addEventListener("error", onError);
    signal?.addEventListener("abort", onAbort, { once: true });
    if (signal?.aborted) {
      finish(null);
      return;
    }
    image.src = src;
  });
}

/**
 * Reads one file into the values an image node needs.
 * Null for a file that is not an image kind we can carry, or one the browser could not
 * decode.
 *
 * `maxWidthPx` is the width an image wider than the page is shrunk to, which the caller
 * reads off the document it is inserting into.
 */
export async function readImageFile(
  file: File,
  maxWidthPx?: number
): Promise<ImageToInsert | null> {
  if (!isImageMime(file.type) || file.size > MAX_IMAGE_BYTES) return null;
  const source = await readDataUrl(file);
  return source === null
    ? null
    : readImageSource(source, { maxWidthPx, alt: null });
}

export interface ImageReadOptions {
  maxWidthPx?: number;
  maxHeightPx?: number;
  preferredExtent?: ImageExtent | null;
  alt?: string | null;
  signal?: AbortSignal;
}

export async function readImageSource(
  source: string,
  {
    maxWidthPx,
    maxHeightPx,
    preferredExtent = null,
    alt = null,
    signal,
  }: ImageReadOptions = {}
): Promise<ImageToInsert | null> {
  const src = boundedImageSrc(source);
  if (src === null) return null;
  const natural = await naturalSize(src, signal);
  const preferred = preferredExtent && {
    width: emuToPx(preferredExtent.cx),
    height: emuToPx(preferredExtent.cy),
  };
  const size = preferred ?? natural;
  const extent =
    natural &&
    size &&
    fittedExtent(size.width, size.height, maxWidthPx, maxHeightPx);
  return extent === null ? null : { src, extent, alt };
}

export async function readImageBlob(
  source: Blob,
  options: ImageReadOptions = {}
): Promise<ImageToInsert | null> {
  if (!isImageMime(source.type) || source.size > MAX_IMAGE_BYTES) return null;
  const dataUrl = await readDataUrl(source);
  return dataUrl === null ? null : readImageSource(dataUrl, options);
}

/**
 * Puts the files into the document one after another, each at the position the one before
 * it left the caret. A file that could not be read is skipped.
 *
 * Each image is fitted to one line of body text on the document's own paper, read once up
 * front: the paper is fixed for as long as the document is open.
 * Reading a file is asynchronous, so the editor may be gone by the time the bytes are
 * there; a torn down view is left alone.
 */
export async function insertImageFiles(
  view: EditorView,
  files: readonly File[]
): Promise<void> {
  const maxWidthPx = documentBodyWidthPx(view.state);
  for (const file of files) {
    const image = await readImageFile(file, maxWidthPx);
    if (!image || view.isDestroyed) continue;
    insertImage(image)(view.state, (tr) => view.dispatch(tr), view);
  }
}

/**
 * The plugin that takes image files pasted or dropped onto the paper.
 *
 * It has to sit ahead of the plain-text plugin, which answers for every paste and drop
 * whether or not it has anything to insert (see the plugin order in `createEditor.ts`).
 * An event carrying no image file is passed on untouched, so text keeps coming in exactly
 * as it did.
 */
export function imageFiles(): Plugin {
  return new Plugin({
    props: {
      handlePaste(view, event) {
        const files = imageFilesIn(event.clipboardData);
        if (files.length === 0) return false;
        void insertImageFiles(view, files);
        return true;
      },
      /**
       * A drag that started inside the editor carries no file, so a picture being moved
       * within the document never reaches this and stays ProseMirror's own move.
       */
      handleDrop(view, event) {
        const files = imageFilesIn(event.dataTransfer);
        if (files.length === 0) return false;
        moveCaretToDrop(view, event);
        void insertImageFiles(view, files);
        return true;
      },
    },
  });
}
