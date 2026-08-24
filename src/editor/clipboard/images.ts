import { MAX_IMAGE_BYTES } from "../../docx/media";
import { isImageMime, toImageExtent, toImageSrc } from "../../ooxml/image";
import { editorClassNames } from "../../styles/classNames";
import {
  boundedImageSrc,
  imageSourceByteLength,
  readImageBlob,
  readImageSource,
} from "../imageFiles";
import type { ImageToInsert } from "../insertImage";

export const PASTED_IMAGE_ATTRIBUTE = "data-docx-pasted-image";

export interface ResolvedClipboardImages {
  source: string;
  images: ReadonlyMap<string, ImageToInsert>;
}

interface ResolveOptions {
  maxWidthPx: number;
  maxHeightPx: number;
  signal: AbortSignal;
}

function extentAttribute(element: HTMLElement) {
  if (!element.classList.contains(editorClassNames.image)) return null;
  try {
    return toImageExtent(
      JSON.parse(element.getAttribute("data-extent") ?? "null")
    );
  } catch {
    return null;
  }
}

function fetchableUrl(source: string): URL | null {
  try {
    const url = new URL(source);
    return ["blob:", "http:", "https:"].includes(url.protocol) ? url : null;
  } catch {
    return null;
  }
}

function canResolve(element: HTMLElement): boolean {
  const source = element.getAttribute("src") ?? "";
  return toImageSrc(source) !== null || fetchableUrl(source) !== null;
}

export function hasResolvableClipboardImages(
  document: Document,
  source: string
): boolean {
  const template = document.createElement("template");
  template.innerHTML = source;
  return Array.from(template.content.querySelectorAll("img")).some(canResolve);
}

function abortIfAsked(signal: AbortSignal): void {
  if (signal.aborted)
    throw new DOMException("Image paste aborted", "AbortError");
}

async function limitedImageBlob(
  response: Response,
  signal: AbortSignal,
  maxBytes: number
): Promise<Blob | null> {
  if (!response.ok) {
    await response.body?.cancel();
    return null;
  }
  const mime = response.headers
    .get("content-type")
    ?.split(";", 1)[0]
    ?.trim()
    .toLowerCase();
  if (!isImageMime(mime)) {
    await response.body?.cancel();
    return null;
  }
  const declaredSize = Number.parseInt(
    response.headers.get("content-length") ?? "",
    10
  );
  if (Number.isFinite(declaredSize) && declaredSize > maxBytes) {
    await response.body?.cancel();
    return null;
  }
  const reader = response.body?.getReader();
  if (!reader) return null;
  const chunks: Uint8Array<ArrayBuffer>[] = [];
  let size = 0;
  try {
    while (true) {
      abortIfAsked(signal);
      const { done, value } = await reader.read();
      if (done) break;
      if (size + value.length > maxBytes) {
        await reader.cancel();
        return null;
      }
      const chunk = Uint8Array.from(value);
      size += chunk.length;
      chunks.push(chunk);
    }
  } catch (error) {
    await reader.cancel().catch(() => undefined);
    throw error;
  }
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  return new Blob([bytes], { type: mime });
}

async function resolveElement(
  element: HTMLElement,
  options: ResolveOptions,
  maxBytes: number
): Promise<ImageToInsert | null> {
  abortIfAsked(options.signal);
  const source = element.getAttribute("src") ?? "";
  const alt = element.getAttribute("alt") || null;
  const readOptions = {
    maxWidthPx: options.maxWidthPx,
    maxHeightPx: options.maxHeightPx,
    preferredExtent: extentAttribute(element),
    alt,
    signal: options.signal,
  };
  if (toImageSrc(source) !== null) {
    const src = boundedImageSrc(source);
    return src === null || imageSourceByteLength(src) > maxBytes
      ? null
      : readImageSource(src, readOptions);
  }
  const url = fetchableUrl(source);
  if (url === null) return null;
  const response = await fetch(url, {
    credentials: "omit",
    referrerPolicy: "no-referrer",
    signal: options.signal,
  });
  const blob = await limitedImageBlob(response, options.signal, maxBytes);
  return blob === null ? null : readImageBlob(blob, readOptions);
}

function replaceWithEditorImage(element: HTMLElement, token: string): void {
  const replacement = element.ownerDocument.createElement("img");
  replacement.setAttribute(PASTED_IMAGE_ATTRIBUTE, token);
  element.replaceWith(replacement);
}

export function resolveClipboardImages(
  document: Document,
  source: string,
  options: ResolveOptions
): Promise<ResolvedClipboardImages> | null {
  const template = document.createElement("template");
  template.innerHTML = source;
  const elements = Array.from(template.content.querySelectorAll("img"));
  for (const element of elements) {
    element.removeAttribute(PASTED_IMAGE_ATTRIBUTE);
  }
  const images = elements.filter(canResolve);
  if (images.length === 0) return null;
  return (async () => {
    let remainingBytes = MAX_IMAGE_BYTES;
    const resolved = new Map<string, ImageToInsert>();
    for (const element of images) {
      abortIfAsked(options.signal);
      let image: ImageToInsert | null = null;
      try {
        image = await resolveElement(element, options, remainingBytes);
      } catch {
        abortIfAsked(options.signal);
        image = null;
      }
      const bytes = image && imageSourceByteLength(image.src);
      if (image === null || bytes === null || bytes > remainingBytes) {
        element.remove();
      } else {
        remainingBytes -= bytes;
        const token = `${resolved.size}`;
        resolved.set(token, image);
        replaceWithEditorImage(element, token);
      }
    }
    return { source: template.innerHTML, images: resolved };
  })();
}
