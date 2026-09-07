/**
 * Where a part sits in the package, asked one way by every reader and writer, and the content
 * types part every part the export adds is declared in.
 *
 * A part beside the body is found through the relationship the main part declares for it, and a
 * part the export adds takes the first name beside the main part that nothing has claimed. Each
 * reader used to find its part with a lookup of its own, so the rule for what counts as the
 * related part lived in five places.
 */

import { elementXml } from "../ooxml/element";
import { DocxExportError } from "../ooxml/errors";
import { rootPrefixOf, splicePart } from "../ooxml/partSplice";
import {
  decodeUtf8,
  elementChildren,
  encodeUtf8,
  parseXml,
} from "../ooxml/xml";
import {
  directoryOf,
  readRelationships,
  relsPathOf,
  resolveTarget,
} from "./relationships";

export const CONTENT_TYPES_PATH = "[Content_Types].xml";

/**
 * Where the part the main part relates under this type sits. null for no such relationship, and
 * for one that points outside the package, which no part of it can stand behind.
 */
export function relatedPartPath(
  parts: ReadonlyMap<string, Uint8Array>,
  mainPartPath: string,
  type: string
): string | null {
  const relationship = readRelationships(parts, relsPathOf(mainPartPath)).find(
    (entry) => entry.type === type && !entry.external
  );
  return relationship === undefined
    ? null
    : resolveTarget(mainPartPath, relationship.target);
}

/** A part as text, and null for a path there is no part at, so a missing relationship reads as a missing part */
export function readPart(
  parts: ReadonlyMap<string, Uint8Array>,
  path: string | null
): string | null {
  const bytes = path === null ? undefined : parts.get(path);
  return bytes ? decodeUtf8(bytes).text : null;
}

/** `comments.xml`, then `comments2.xml`, and so on beside the main part: the first name no part of the package has taken */
export function availablePartPath(
  parts: ReadonlyMap<string, Uint8Array>,
  mainPartPath: string,
  stem: string
): string {
  const directory = directoryOf(mainPartPath);
  const taken = new Set(Array.from(parts.keys(), (path) => path.toLowerCase()));
  for (let suffix = 0; ; suffix += 1) {
    const path = `${directory}${stem}${suffix === 0 ? "" : suffix + 1}.xml`;
    if (!taken.has(path.toLowerCase())) return path;
  }
}

/**
 * The content types part as an export leaves it: what the package declared when it was opened,
 * and every declaration the planners asked for on top, written once at the end.
 *
 * Each writer that added a part used to splice its own declaration in and hand the result to the
 * next, so what the part came to hold depended on the order they were called in, and a writer
 * called without the last one's result wrote over its declaration.
 */
export interface ContentTypeWriter {
  /** Declares the type of the one part at this path */
  addOverride(partPath: string, contentType: string): void;
  /** Declares the type of every part with this extension */
  addDefault(extension: string, contentType: string): void;
  /**
   * The part with everything added, or null when nothing was added or everything asked for was
   * declared already. The additions go right after the opening tag, the defaults ahead of the
   * overrides, whichever order they were asked in.
   */
  part(): Uint8Array | null;
}

/**
 * What the part declares, each name lowercased: a part name and an extension are both matched
 * without regard to case (ECMA-376 Part 2, 9.1.2.2).
 *
 * Only a `Default` covers an extension and only an `Override` covers a part name. An override
 * pins down a part a producer already had, so it says nothing about a media part that is not
 * there yet.
 */
function declaredIn(text: string): {
  defaults: Set<string>;
  overrides: Set<string>;
  xmlDefaults: Set<string>;
  xmlOverrides: Set<string>;
} {
  let root: Element;
  try {
    root = parseXml(text).documentElement;
  } catch (cause) {
    throw new DocxExportError(
      "malformed-xml",
      `${CONTENT_TYPES_PATH} could not be parsed`,
      { cause }
    );
  }
  const defaults = new Set<string>();
  const overrides = new Set<string>();
  const xmlDefaults = new Set<string>();
  const xmlOverrides = new Set<string>();
  for (const el of elementChildren(root)) {
    const extension = el.getAttribute("Extension");
    const partName = el.getAttribute("PartName");
    const contentType = (el.getAttribute("ContentType") ?? "")
      .split(";")[0]
      .trim()
      .toLowerCase();
    const xml =
      contentType === "application/xml" ||
      contentType === "text/xml" ||
      contentType.endsWith("+xml");
    if (el.localName === "Default" && extension !== null) {
      defaults.add(extension.toLowerCase());
      if (xml) xmlDefaults.add(extension.toLowerCase());
    }
    if (el.localName === "Override" && partName !== null) {
      overrides.add(partName.toLowerCase());
      if (xml) xmlOverrides.add(partName.toLowerCase());
    }
  }
  return { defaults, overrides, xmlDefaults, xmlOverrides };
}

/** XML parts named by OPC content types, even when their URI has no .xml suffix. */
export function declaredXmlParts(
  types: Uint8Array,
  paths: Iterable<string>
): ReadonlySet<string> {
  const declared = declaredIn(decodeUtf8(types).text);
  return new Set(
    Array.from(paths).filter((path) => {
      const name = `/${path}`.toLowerCase();
      if (declared.overrides.has(name)) return declared.xmlOverrides.has(name);
      const extension = path.slice(path.lastIndexOf(".") + 1).toLowerCase();
      return declared.xmlDefaults.has(extension);
    })
  );
}

export function contentTypeWriter(
  parts: ReadonlyMap<string, Uint8Array>
): ContentTypeWriter {
  const overrides = new Map<string, readonly [string, string]>();
  const defaults = new Map<string, readonly [string, string]>();
  return {
    addOverride: (partPath, contentType) => {
      const key = partPath.toLowerCase();
      if (!overrides.has(key)) overrides.set(key, [partPath, contentType]);
    },
    addDefault: (extension, contentType) => {
      const key = extension.toLowerCase();
      if (!defaults.has(key)) defaults.set(key, [extension, contentType]);
    },
    part: () => {
      if (overrides.size === 0 && defaults.size === 0) return null;
      const original = parts.get(CONTENT_TYPES_PATH);
      if (!original) {
        // Writing this part from scratch would mean guessing the type of every other part in
        // the package, so the export stops rather than hand back a file Word refuses to open
        throw new DocxExportError(
          "missing-content-types",
          `cannot declare a part in a package that has no ${CONTENT_TYPES_PATH}`
        );
      }
      const { text, hadBom } = decodeUtf8(original);
      const declared = declaredIn(text);
      const prefix = rootPrefixOf(text);
      const declarations = [
        ...Array.from(defaults.values())
          .filter(
            ([extension]) => !declared.defaults.has(extension.toLowerCase())
          )
          .map(([extension, contentType]) =>
            elementXml(`${prefix}Default`, [
              ["Extension", extension],
              ["ContentType", contentType],
            ])
          ),
        ...Array.from(overrides.values())
          .filter(
            ([partPath]) =>
              !declared.overrides.has(`/${partPath}`.toLowerCase())
          )
          .map(([partPath, contentType]) =>
            elementXml(`${prefix}Override`, [
              ["PartName", `/${partPath}`],
              ["ContentType", contentType],
            ])
          ),
      ];
      if (declarations.length === 0) return null;
      return encodeUtf8(
        splicePart(text, { root: "Types", prepend: declarations.join("") }),
        hadBom
      );
    },
  };
}
