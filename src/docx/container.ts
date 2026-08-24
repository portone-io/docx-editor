import { inflateSync, zipSync } from "fflate";
import { DocxImportError } from "../ooxml/errors";

/**
 * Caps on one part and on the whole package. The bytes are untrusted and a zip chooses its own
 * inflated size (a few hundred KB of zeros reach hundreds of MB), so both the stored size and the
 * declared inflated size are checked off the directory before any entry is inflated. The caps sit
 * far above a real document (low megabytes), so nothing legitimate is refused.
 */
export const MAX_PART_BYTES = 32 * 1024 * 1024;
export const MAX_PACKAGE_BYTES = 64 * 1024 * 1024;

/**
 * Whether the entry name reaches outside the package.
 *
 * An entry name travels into the exported zip as it came, so a name like `../../evil`
 * would hand whatever unpacks that file afterwards a path outside the directory it
 * unpacks into. A backslash counts as a separator here, since Windows extractors read it
 * as one. OPC part names are relative and carry no `..` segment, so no real docx has one.
 */
function escapesPackage(name: string): boolean {
  const segments = name.split(/[/\\]/);
  return (
    segments[0] === "" ||
    /^[A-Za-z]:$/.test(segments[0]) ||
    segments.includes("..")
  );
}

const END_OF_DIRECTORY = 0x06054b50;
const DIRECTORY_ENTRY = 0x02014b50;
const LOCAL_HEADER = 0x04034b50;

/** The end record is 22 bytes and the comment behind it may run to 65535 */
const END_RECORD_SEARCH = 22 + 0xffff;

const STORED = 0;
const DEFLATED = 8;

/** OPC part names are ASCII, which UTF-8 decodes safely. */
const nameDecoder = new TextDecoder("utf-8");

/** Every unreadable-package path is the same refusal: not a docx. */
function unreadable(why: string): never {
  throw new DocxImportError("not-a-docx", why);
}

interface Entry {
  name: string;
  compression: number;
  storedSize: number;
  /** Size it claims to inflate to. */
  declaredSize: number;
  /** Data offset, taken from the entry's own local header. */
  dataAt: number;
}

/** End-of-central-directory record, searched from the tail. */
function endOfDirectory(view: DataView): number {
  const earliest = Math.max(0, view.byteLength - END_RECORD_SEARCH);
  for (let at = view.byteLength - 22; at >= earliest; at -= 1) {
    if (view.getUint32(at, true) === END_OF_DIRECTORY) return at;
  }
  return unreadable("the bytes are not a zip package");
}

/**
 * Start of an entry's data. The local header's own name and extra-field lengths are used, not the
 * directory's: a writer may give the two different extra fields.
 */
function dataStart(view: DataView, headerAt: number): number {
  if (
    headerAt + 30 > view.byteLength ||
    view.getUint32(headerAt, true) !== LOCAL_HEADER
  ) {
    return unreadable("an entry of the package has no readable header");
  }
  const nameLength = view.getUint16(headerAt + 26, true);
  const extraLength = view.getUint16(headerAt + 28, true);
  return headerAt + 30 + nameLength + extraLength;
}

/**
 * Every entry, read from the central directory rather than the local headers: a local header may
 * omit sizes (a streaming writer only learns them after the data), while the directory always has them.
 */
function directoryEntries(bytes: Uint8Array): Entry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const end = endOfDirectory(view);
  const count = view.getUint16(end + 10, true);
  const entries: Entry[] = [];
  let at = view.getUint32(end + 16, true);
  for (let read = 0; read < count; read += 1) {
    if (
      at + 46 > view.byteLength ||
      view.getUint32(at, true) !== DIRECTORY_ENTRY
    ) {
      return unreadable("the package's directory does not read as one");
    }
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    entries.push({
      name: nameDecoder.decode(bytes.subarray(at + 46, at + 46 + nameLength)),
      compression: view.getUint16(at + 10, true),
      storedSize: view.getUint32(at + 20, true),
      declaredSize: view.getUint32(at + 24, true),
      dataAt: dataStart(view, view.getUint32(at + 42, true)),
    });
    at += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

/**
 * An entry's bytes, refused if they are not the size it declares. The output buffer is one byte
 * longer than declared because fflate silently drops writes past the buffer end: an entry
 * understating its size would otherwise inflate truncated and be repacked as-is. The extra byte
 * tells a part that inflates to exactly its declared size from one that inflates to more.
 */
function entryBytes(bytes: Uint8Array, entry: Entry): Uint8Array {
  const end = entry.dataAt + entry.storedSize;
  if (end > bytes.length) {
    return unreadable("an entry of the package reaches past its end");
  }
  const held = bytes.subarray(entry.dataAt, end);
  if (entry.storedSize === 0) {
    if (entry.declaredSize !== 0) {
      return unreadable("an empty entry of the package declares a size");
    }
    return held;
  }
  if (entry.compression === STORED) {
    if (entry.storedSize !== entry.declaredSize) {
      return unreadable("a stored entry does not hold the size it declares");
    }
    return held.slice();
  }
  if (entry.compression !== DEFLATED) {
    return unreadable(
      "the package holds an entry compressed in a way we cannot read"
    );
  }
  const inflated = inflateSync(held, {
    out: new Uint8Array(entry.declaredSize + 1),
  });
  if (inflated.length !== entry.declaredSize) {
    return unreadable("an entry does not inflate to the size it declares");
  }
  return inflated;
}

/** Collects every part in the zip, keeping the original order */
export function openParts(bytes: Uint8Array): Map<string, Uint8Array> {
  const parts = new Map<string, Uint8Array>();
  let packageBytes = 0;
  try {
    for (const entry of directoryEntries(bytes)) {
      if (escapesPackage(entry.name)) {
        unreadable("the package holds an entry whose name reaches outside it");
      }
      // A stored entry is never inflated, so it costs the larger of the two sizes
      packageBytes += Math.max(entry.storedSize, entry.declaredSize);
      if (
        entry.storedSize > MAX_PART_BYTES ||
        entry.declaredSize > MAX_PART_BYTES ||
        packageBytes > MAX_PACKAGE_BYTES
      ) {
        throw new DocxImportError(
          "too-large",
          "the docx file asks to inflate to more than we open"
        );
      }
      parts.set(entry.name, entryBytes(bytes, entry));
    }
    return parts;
  } catch (error) {
    // A refusal of our own carries the reason already; anything else is a file that is not a docx, or a corrupted one
    if (error instanceof DocxImportError) throw error;
    throw new DocxImportError("not-a-docx", "could not open the docx file", {
      cause: error,
    });
  }
}

/** Builds a zip that swaps in only the rewritten parts and keeps every other part's original bytes as they were */
export function repackParts(
  parts: Map<string, Uint8Array>,
  replacements: Map<string, Uint8Array>
): Uint8Array {
  const zippable: Record<string, Uint8Array> = {};
  for (const [name, data] of parts) {
    zippable[name] = replacements.get(name) ?? data;
  }
  for (const [name, data] of replacements) {
    if (!parts.has(name)) zippable[name] = data;
  }
  // Pin the timestamp so that exporting the same document twice yields the same file
  return zipSync(zippable, { level: 6, mtime: new Date(1980, 0, 1) });
}
