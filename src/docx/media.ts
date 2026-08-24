/**
 * Reads and writes package image parts. Nodes keep data URLs so undoable document states own their
 * bytes without blob-URL lifetime management; `MAX_IMAGE_BYTES` bounds the memory tradeoff.
 */

import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import {
  type ImageMime,
  imageBase64Of,
  imageMimeOf,
  toImageSrc,
} from "../ooxml/image";
import { decodeUtf8, encodeUtf8, R_NS } from "../ooxml/xml";
import {
  directoryOf,
  type Relationship,
  type RelationshipWriter,
  readRelationships,
  relsPathOf,
  resolveTarget,
} from "./relationships";
import type { SessionStore } from "./session";

const IMAGE_REL_TYPE = `${R_NS}/image`;

const CONTENT_TYPES_PATH = "[Content_Types].xml";

/** The extension a media part gets, per kind we can write */
const EXTENSION_BY_MIME: Record<ImageMime, string> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/gif": "gif",
  "image/bmp": "bmp",
};

/** The kind a media part holds, per extension we can read */
const MIME_BY_EXTENSION = new Map<string, ImageMime>([
  ["png", "image/png"],
  ["jpeg", "image/jpeg"],
  ["jpg", "image/jpeg"],
  ["jpe", "image/jpeg"],
  ["gif", "image/gif"],
  ["bmp", "image/bmp"],
]);

/**
 * Chunked, because `String.fromCharCode` takes its arguments one at a time and a whole
 * image would overrun the argument limit
 */
function toBase64(bytes: Uint8Array): string {
  const CHUNK = 0x8000;
  let binary = "";
  for (let at = 0; at < bytes.length; at += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(at, at + CHUNK));
  }
  return btoa(binary);
}

function fromBase64(base64: string): Uint8Array {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let at = 0; at < binary.length; at += 1) {
    bytes[at] = binary.charCodeAt(at);
  }
  return bytes;
}

function dataUrl(bytes: Uint8Array, mime: ImageMime): string {
  return `data:${mime};base64,${toBase64(bytes)}`;
}

function mimeOfPath(path: string): ImageMime | undefined {
  const dot = path.lastIndexOf(".");
  return dot === -1
    ? undefined
    : MIME_BY_EXTENSION.get(path.slice(dot + 1).toLowerCase());
}

/** The images the body can refer to, as data URLs keyed by relationship id */
export type ImageSources = ReadonlyMap<string, string>;

export const NO_IMAGES: ImageSources = new Map();

/**
 * The most one image may weigh before it is left undrawn.
 *
 * The document state holds every image as base64 for as long as the editor is open, and
 * base64 costs a third more than the bytes, so an untrusted document would otherwise
 * decide by itself how much memory the session takes. An image past the cap takes the same
 * road as one no browser can draw rather than the document being refused: nothing is lost,
 * since the drawing and its media part still ride back out untouched.
 * No photograph a contract carries comes near this.
 */
export const MAX_IMAGE_BYTES = 16 * 1024 * 1024;

/**
 * Every image the body part has a relationship to.
 *
 * A relationship we cannot follow (a linked file, a missing part, a kind no browser
 * draws, an image past `MAX_IMAGE_BYTES`) is simply left out, and the drawing that points
 * at it stays on the preservation path.
 */
export function readImageSources(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): ImageSources {
  const sources = new Map<string, string>();
  for (const rel of readRelationships(parts, relsPathOf(mainPartPath))) {
    if (rel.type !== IMAGE_REL_TYPE || rel.external) continue;
    const mime = mimeOfPath(rel.target);
    const bytes = parts.get(resolveTarget(mainPartPath, rel.target));
    if (!mime || !bytes || bytes.length > MAX_IMAGE_BYTES) continue;
    sources.set(rel.id, dataUrl(bytes, mime));
  }
  return sources;
}

/**
 * A fingerprint of the bytes, so the same image inserted twice becomes one media part.
 *
 * Two independent FNV-1a passes plus the length. This is only ever used to recognize
 * bytes we have already written, never to prove anything about them.
 */
function contentKey(bytes: Uint8Array): string {
  let a = 0x811c9dc5;
  let b = 0xc2b2ae35;
  for (const byte of bytes) {
    a = Math.imul(a ^ byte, 0x01000193) >>> 0;
    b = Math.imul(b ^ byte, 0x85ebca6b) >>> 0;
  }
  const hex = (value: number) => value.toString(16).padStart(8, "0");
  return `${bytes.length.toString(36)}${hex(a)}${hex(b)}`;
}

/**
 * What the body needs to know while it is being written out: which relationship each
 * inserted image ended up on, and a drawing id no other drawing in the document holds.
 *
 * The ids are handed out one at a time rather than planned per image, because the same
 * image inserted twice shares one media part but is still two drawings, and Word expects
 * every drawing in a document to have an id of its own.
 */
export interface ImageRefs {
  relIdOf(src: string): string | undefined;
  takeDocPrId(): number;
}

function imageRefs(
  relIdBySrc: ReadonlyMap<string, string>,
  firstDocPrId: number
): ImageRefs {
  let nextDocPrId = firstDocPrId;
  return {
    relIdOf: (src) => relIdBySrc.get(src),
    takeDocPrId: () => {
      const id = nextDocPrId;
      nextDocPrId += 1;
      return id;
    },
  };
}

/** What a serializer called on its own, outside an export, has to work with */
export const NO_IMAGE_REFS: ImageRefs = {
  relIdOf: () => undefined,
  // Never reached: an image with no relationship to point at is refused before it is drawn
  takeDocPrId: () => 1,
};

export interface MediaAdditions {
  refs: ImageRefs;
  /** The parts to write into the zip alongside the body */
  parts: ReadonlyMap<string, Uint8Array>;
}

/** The srcs of the images that were not in the document when it was opened, in document order */
function insertedImageSrcs(doc: PMNode): string[] {
  const srcs: string[] = [];
  doc.descendants((node) => {
    // An imported image holds its original XML and needs nothing added to the package
    if (node.type.name !== "image" || typeof node.attrs.xml === "string") {
      return true;
    }
    const src = toImageSrc(node.attrs.src);
    if (src !== null && !srcs.includes(src)) srcs.push(src);
    return true;
  });
  return srcs;
}

interface InsertedImage {
  bytes: Uint8Array;
  extension: string;
  key: string;
}

function decodeImage(src: string): InsertedImage {
  const mime = imageMimeOf(src);
  if (mime === null) {
    throw new DocxExportError(
      "unsupported-content",
      "an inserted image is of a kind we cannot write"
    );
  }
  try {
    const bytes = fromBase64(imageBase64Of(src));
    return {
      bytes,
      extension: EXTENSION_BY_MIME[mime],
      key: contentKey(bytes),
    };
  } catch (error) {
    throw new DocxExportError(
      "unsupported-content",
      "an inserted image's bytes could not be decoded",
      { cause: error }
    );
  }
}

/** The relationship id of every image already in the package, keyed by a fingerprint of its bytes */
function imageRelIdsByContent(
  session: SessionStore,
  rels: readonly Relationship[]
): Map<string, string> {
  const byContent = new Map<string, string>();
  for (const rel of rels) {
    if (rel.type !== IMAGE_REL_TYPE || rel.external) continue;
    const bytes = session.parts.get(
      resolveTarget(session.mainPartPath, rel.target)
    );
    if (!bytes) continue;
    const key = contentKey(bytes);
    if (!byContent.has(key)) byContent.set(key, rel.id);
  }
  return byContent;
}

const DOC_PR_ID = /<(?:[\w.-]+:)?docPr\b[^>]*\bid="(\d+)"/g;

/** The largest drawing id the document already uses, so a new drawing does not reuse one */
function maxDocPrId(session: SessionStore): number {
  let max = 0;
  for (const block of session.blocks) {
    for (const match of block.xml.matchAll(DOC_PR_ID)) {
      max = Math.max(max, Number.parseInt(match[1], 10));
    }
  }
  return max;
}

const TYPES_OPEN_TAG = /<(?:[\w.-]+:)?Types\b[^>]*>/;

/**
 * The content types part with a `Default` added for every extension it does not declare
 * yet. null when it already declares them all.
 *
 * Only the defaults are looked at: an `Override` naming one single part is how a producer
 * pins down a part we are not adding, so it cannot cover a media part that is not there
 * yet.
 */
function withImageContentTypes(
  parts: Map<string, Uint8Array>,
  extensions: ReadonlySet<string>
): Uint8Array | null {
  const original = parts.get(CONTENT_TYPES_PATH);
  if (!original) {
    // Writing this part from scratch would mean guessing the type of every other part in
    // the package. We stop instead of handing back a file Word refuses to open.
    throw new DocxExportError(
      "missing-content-types",
      `cannot add an image to a package that has no ${CONTENT_TYPES_PATH}`
    );
  }
  const { text, hadBom } = decodeUtf8(original);
  const missing = Array.from(extensions).filter(
    (extension) =>
      !new RegExp(`<Default[^>]+Extension="${extension}"`, "i").test(text)
  );
  if (missing.length === 0) return null;

  const open = TYPES_OPEN_TAG.exec(text);
  if (!open) {
    throw new DocxExportError(
      "malformed-xml",
      `${CONTENT_TYPES_PATH} has no Types element`
    );
  }
  const declarations = missing
    .map((extension) => {
      const mime = MIME_BY_EXTENSION.get(extension);
      return `<Default Extension="${extension}" ContentType="${mime}"/>`;
    })
    .join("");
  const at = open.index + open[0].length;
  return encodeUtf8(text.slice(0, at) + declarations + text.slice(at), hadBom);
}

/**
 * What the package needs for the images that were inserted during editing: the media parts,
 * the relationships pointing at them, and the content types declaring them.
 * null when every image in the document came out of the document itself.
 *
 * The relationships go through the writer the export shares between its writers
 * (`docx/relationships`), so the part is written once with everything in it.
 *
 * Bytes that are already in the package are not written a second time: an inserted image
 * whose fingerprint matches one that is already there reuses its relationship.
 */
export function planImageMedia(
  doc: PMNode,
  session: SessionStore,
  relationships: RelationshipWriter
): MediaAdditions | null {
  const srcs = insertedImageSrcs(doc);
  if (srcs.length === 0) return null;

  const relIdByContent = imageRelIdsByContent(session, relationships.opened);
  const directory = directoryOf(session.mainPartPath);

  const relIdBySrc = new Map<string, string>();
  const parts = new Map<string, Uint8Array>();
  const extensions = new Set<string>();

  for (const src of srcs) {
    const image = decodeImage(src);
    const reused = relIdByContent.get(image.key);
    if (reused !== undefined) {
      relIdBySrc.set(src, reused);
      continue;
    }

    const target = `media/image-${image.key}.${image.extension}`;
    const relId = relationships.add({ type: IMAGE_REL_TYPE, target });
    parts.set(directory + target, image.bytes);
    extensions.add(image.extension);
    relIdByContent.set(image.key, relId);
    relIdBySrc.set(src, relId);
  }

  if (extensions.size > 0) {
    const contentTypes = withImageContentTypes(session.parts, extensions);
    if (contentTypes) parts.set(CONTENT_TYPES_PATH, contentTypes);
  }
  return { refs: imageRefs(relIdBySrc, maxDocPrId(session) + 1), parts };
}
