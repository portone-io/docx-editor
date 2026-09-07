/**
 * Turns the edited document back into docx bytes.
 *
 * The body (document.xml) is always rebuilt from preserved and edited blocks. numbering.xml is
 * rewritten only when a list was newly started, media parts only when an image was inserted, and
 * the Comments part only when comments changed. Content types and relationships change only when
 * one of those additions needs them.
 *
 * The body is written before that relationships part, because a link asks for its relationship as
 * it is written (`docx/hyperlink`), and both writers hand out ids through the one writer so that
 * they cannot pick the same one.
 */

import type { Node as PMNode } from "prosemirror-model";
import { addListDefinitions } from "../numbering/writeNumbering";
import { DocxExportError } from "../ooxml/errors";
import {
  decodeUtf8,
  encodeUtf8,
  parseXml,
  W_NS,
  withXmlParser,
  type XmlParser,
} from "../ooxml/xml";
import { sameSource } from "../schema/sourceEquality";
import { planCommentParts } from "./comments";
import { repackParts } from "./container";
import type { ExportRefs } from "./exportRefs";
import {
  type FidelityCollector,
  type FidelityNote,
  fidelityNotesOf,
} from "./fidelity";
import { hyperlinkRefs } from "./hyperlink";
import { withUniqueIdentities } from "./identities";
import { problemsOf } from "./invariants";
import { NO_IMAGE_REFS, planImageMedia } from "./media";
import { newNumIds, numberingPartOf } from "./newLists";
import {
  readRelationships,
  relationshipWriter,
  relsPathOf,
} from "./relationships";
import { serializeBlock } from "./serializeBlock";
import {
  type DocxSession,
  originalBlock,
  type SessionStore,
  sessionOf,
} from "./session";

/**
 * An unchanged block is exported with its original XML as is; only a changed block is rebuilt.
 *
 * Unchanged is judged by `sameSource` rather than by `Node.eq`, because opening a file works the
 * display attrs out again (`schema/attrRoles`) and a block rebuilt over that would lose the markup
 * the writer does not model, the properties of a cell continuing a vertical merge among it.
 */
function blockXml(
  node: PMNode,
  session: SessionStore,
  refs: ExportRefs
): string {
  const imported = originalBlock(node, session);
  if (imported && sameSource(node, imported.node)) return imported.xml;
  return serializeBlock(node, session, refs);
}

function buildDocumentXml(
  doc: PMNode,
  session: SessionStore,
  refs: ExportRefs
): string {
  const pieces: string[] = [session.documentPrefix];
  withUniqueIdentities(doc).forEach((child) => {
    pieces.push(blockXml(child, session, refs));
  });
  pieces.push(session.documentSuffix);
  return pieces.join("");
}

function localAttribute(el: Element, name: string): string | null {
  return (
    Array.from(el.attributes).find((attribute) => attribute.localName === name)
      ?.value ?? null
  );
}

/** OOXML requires every bookmark end to identify an earlier unmatched start. */
function assertBookmarkPairs(documentXml: string): void {
  let root: Element;
  try {
    root = parseXml(documentXml).documentElement;
  } catch (cause) {
    throw new DocxExportError(
      "malformed-xml",
      "the exported main document XML could not be parsed",
      { cause }
    );
  }
  const open = new Set<string>();
  const used = new Set<string>();
  for (const element of root.getElementsByTagName("*")) {
    if (element.namespaceURI !== W_NS) continue;
    if (
      element.localName !== "bookmarkStart" &&
      element.localName !== "bookmarkEnd"
    ) {
      continue;
    }
    const id = localAttribute(element, "id");
    if (id === null) {
      throw new DocxExportError(
        "malformed-xml",
        `a ${element.localName} has no id`
      );
    }
    if (element.localName === "bookmarkStart") {
      if (used.has(id)) {
        throw new DocxExportError(
          "malformed-xml",
          `bookmark ${id} has more than one start marker`
        );
      }
      open.add(id);
      used.add(id);
    } else if (!open.delete(id)) {
      throw new DocxExportError(
        "malformed-xml",
        `bookmark ${id} ends without an earlier start marker`
      );
    }
  }
  const missingEnd = open.values().next().value;
  if (missingEnd !== undefined) {
    throw new DocxExportError(
      "malformed-xml",
      `bookmark ${missingEnd} has no end marker`
    );
  }
}

/**
 * A rewritten numbering.xml, produced only when a new list appeared.
 * The original text is left as is and only the new definitions are spliced in. null if there is no new list.
 */
function newNumberingPart(
  doc: PMNode,
  session: SessionStore
): { path: string; bytes: Uint8Array } | null {
  const added = newNumIds(doc, session);
  const original = numberingPartOf(session);
  // A new list in a document with no numbering.xml is refused by the invariant list before
  // anything is written, so a missing part here has nothing to hold
  if (added.length === 0 || original === null) return null;

  const { text, hadBom } = decodeUtf8(original.bytes);
  return {
    path: original.path,
    bytes: encodeUtf8(addListDefinitions(text, added), hadBom),
  };
}

/** What a caller may say about a write beyond handing over the document and its session */
export interface ExportOptions {
  /**
   * The parser the original XML kept in the session is read back through. Left out, the
   * `DOMParser` global is used, and a runtime carrying none refuses the write with
   * `no-xml-parser`.
   */
  xmlParser?: XmlParser;
}

export function exportDocx(
  doc: PMNode,
  session: DocxSession,
  options?: ExportOptions
): Uint8Array {
  return exportDocxReport(doc, session, options).bytes;
}

/**
 * The file to write, and what writing it could not carry across as it stood.
 *
 * The notes are the ones `documentFidelity` reads off the same document, followed by whatever the
 * writer had to approximate on the way out. A host that hands a file to somebody else can say what
 * that file no longer holds without opening it again.
 *
 * The first problem `exportProblems` reports is thrown before anything is written, so a refusal a
 * caller could have asked about ahead of time arrives with the same code and message it would
 * have read there.
 */
export function exportDocxReport(
  doc: PMNode,
  session: DocxSession,
  options?: ExportOptions
): { bytes: Uint8Array; notes: FidelityNote[] } {
  return withXmlParser(options?.xmlParser, () => {
    const store = sessionOf(session);
    const problem = problemsOf(doc, store)[0];
    if (problem) throw new DocxExportError(problem.code, problem.message);
    const approximated: FidelityNote[] = [];
    const bytes = writeDocx(doc, store, {
      add: (note) => approximated.push(note),
    });
    return {
      bytes,
      notes: [...fidelityNotesOf(doc, store.mainPartPath), ...approximated],
    };
  });
}

function writeDocx(
  doc: PMNode,
  store: SessionStore,
  notes: FidelityCollector
): Uint8Array {
  const relsPath = relsPathOf(store.mainPartPath);
  const relationships = relationshipWriter(
    readRelationships(store.parts, relsPath)
  );
  // The body has to know which relationship a newly inserted image ends up on, so the
  // media is planned before the body is written
  const media = planImageMedia(doc, store, relationships);
  const comments = planCommentParts(
    doc,
    store,
    relationships,
    media?.parts.get("[Content_Types].xml")
  );
  const documentXml = buildDocumentXml(doc, store, {
    images: media?.refs ?? NO_IMAGE_REFS,
    links: hyperlinkRefs(relationships),
    notes,
  });
  assertBookmarkPairs(documentXml);

  const replacements = new Map<string, Uint8Array>([
    [store.mainPartPath, encodeUtf8(documentXml, store.documentHadBom)],
  ]);
  const numbering = newNumberingPart(doc, store);
  if (numbering) replacements.set(numbering.path, numbering.bytes);
  for (const [path, bytes] of media?.parts ?? []) {
    replacements.set(path, bytes);
  }
  for (const [path, bytes] of comments?.parts ?? []) {
    replacements.set(path, bytes);
  }
  const rels = relationships.part(store.parts.get(relsPath));
  if (rels) replacements.set(relsPath, rels);
  return repackParts(store.parts, replacements);
}
