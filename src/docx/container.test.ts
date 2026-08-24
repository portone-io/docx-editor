import { zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { importErrorCode } from "../__testing__/docx";
import { MAX_PACKAGE_BYTES, MAX_PART_BYTES, openParts } from "./container";

const encode = (text: string) => new TextEncoder().encode(text);

/**
 * A part that is nothing but zeros: it deflates to a few kilobytes, so the zip stays tiny
 * while the size it asks for once opened is the whole point of the test
 */
const zeros = (bytes: number) => new Uint8Array(bytes);

/** Where a little-endian field sits, searched for by the record it belongs to */
function fieldAt(zip: Uint8Array, signature: number, offset: number): number {
  const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
  for (let at = 0; at + 4 <= zip.length; at += 1) {
    if (view.getUint32(at, true) === signature) return at + offset;
  }
  throw new Error("the zip holds no such record");
}

/**
 * The same package with the size its one entry says it inflates to rewritten, in both the
 * directory and the entry's own header, the way a crafted or a damaged file carries it.
 *
 * A size read off a header is a size the file is free to misstate, which is what every entry
 * below does.
 */
function understating(zip: Uint8Array, declared: number): Uint8Array {
  const patched = zip.slice();
  const view = new DataView(patched.buffer);
  view.setUint32(fieldAt(patched, 0x04034b50, 22), declared, true);
  view.setUint32(fieldAt(patched, 0x02014b50, 24), declared, true);
  return patched;
}

describe("opening a package", () => {
  it("reads every part, keeping the order they came in", () => {
    const parts = openParts(
      zipSync({
        "_rels/.rels": encode("<Relationships/>"),
        "word/document.xml": encode("<w:document/>"),
      })
    );
    expect(Array.from(parts.keys())).toEqual([
      "_rels/.rels",
      "word/document.xml",
    ]);
  });

  it("refuses bytes that are not a zip at all", () => {
    expect(importErrorCode(() => openParts(encode("not a zip")))).toBe(
      "not-a-docx"
    );
  });
});

/**
 * The bytes come from outside, so the work opening them asks for has to be bounded before
 * any of it is done: a zip of a few kilobytes can ask for hundreds of megabytes
 */
describe("the size a package may inflate to", () => {
  it("opens a package that stays inside both caps", () => {
    const parts = openParts(zipSync({ "word/document.xml": zeros(1024) }));
    expect(parts.get("word/document.xml")?.length).toBe(1024);
  });

  it("refuses a single part that inflates past the part cap", () => {
    const bomb = zipSync({ "word/document.xml": zeros(MAX_PART_BYTES + 1) });
    // The refusal costs a fraction of what opening the part would have
    expect(bomb.length).toBeLessThan(1024 * 1024);
    expect(importErrorCode(() => openParts(bomb))).toBe("too-large");
  });

  /**
   * A stored entry is handed out as the bytes it holds, and nothing ever inflates it, so the
   * size the caps have to bound is the one the entry takes rather than the one it declares
   */
  it("refuses a stored part whose bytes pass the part cap", () => {
    const stored = zipSync(
      { "word/document.xml": zeros(MAX_PART_BYTES + 1) },
      { level: 0 }
    );
    expect(importErrorCode(() => openParts(stored))).toBe("too-large");
    expect(importErrorCode(() => openParts(understating(stored, 1024)))).toBe(
      "too-large"
    );
  });

  it("refuses parts that stay under the part cap but pass the package cap together", () => {
    const part = zeros(Math.floor(MAX_PACKAGE_BYTES / 3) + 1);
    expect(part.length).toBeLessThan(MAX_PART_BYTES);
    const bomb = zipSync({
      "word/document.xml": part,
      "word/styles.xml": part,
      "word/numbering.xml": part,
    });
    expect(importErrorCode(() => openParts(bomb))).toBe("too-large");
  });
});

/**
 * An entry understating its size is what no cap can bound on its own: the inflated bytes go
 * into a buffer of exactly the size the header claims, and whatever runs past the end of that
 * buffer is dropped rather than refused, so the part would come back quietly truncated and
 * ride back out of the export that way.
 */
describe("a part that does not hold what it says it does", () => {
  const document = encode("<w:document>a part of some size</w:document>");

  it("is refused rather than opened truncated", () => {
    const crafted = understating(zipSync({ "word/document.xml": document }), 4);
    expect(importErrorCode(() => openParts(crafted))).toBe("not-a-docx");
  });

  it("is refused where it overstates its size as well", () => {
    const crafted = understating(
      zipSync({ "word/document.xml": document }),
      document.length + 8
    );
    expect(importErrorCode(() => openParts(crafted))).toBe("not-a-docx");
  });

  it("is refused where it is stored rather than deflated", () => {
    const crafted = understating(
      zipSync({ "word/document.xml": document }, { level: 0 }),
      4
    );
    expect(importErrorCode(() => openParts(crafted))).toBe("not-a-docx");
  });

  it("opens the same package untouched", () => {
    const parts = openParts(zipSync({ "word/document.xml": document }));
    expect(parts.get("word/document.xml")).toEqual(document);
  });
});

/**
 * A part name travels into the exported zip as it came, so a name reaching outside the
 * package would hand whatever unpacks that file afterwards a path of its own choosing
 */
describe("entry names that reach outside the package", () => {
  const escaping = [
    "../../evil.xml",
    "word/../../evil.xml",
    "..\\..\\evil.xml",
    "/etc/passwd",
    "C:\\evil.xml",
  ];

  it.each(escaping)("refuses a package holding %s", (name) => {
    const traversal = zipSync({
      "word/document.xml": encode("<w:document/>"),
      [name]: encode("owned"),
    });
    expect(importErrorCode(() => openParts(traversal))).toBe("not-a-docx");
  });

  it("keeps the names a real package uses", () => {
    const parts = openParts(
      zipSync({
        "[Content_Types].xml": encode("<Types/>"),
        "word/_rels/document.xml.rels": encode("<Relationships/>"),
        "customXml/item1.xml": encode("<item/>"),
      })
    );
    expect(Array.from(parts.keys())).toEqual([
      "[Content_Types].xml",
      "word/_rels/document.xml.rels",
      "customXml/item1.xml",
    ]);
  });
});
