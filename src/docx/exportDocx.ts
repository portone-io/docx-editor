/**
 * Turns the edited document back into docx bytes.
 *
 * The body (document.xml) is always rebuilt from preserved and edited blocks. Every other part is
 * written by a planner (`./partPlan`) that answers only when the document gives it something to
 * write: numbering.xml when a list was newly started, the comment parts when comments changed,
 * media parts when an image was inserted. Content types and relationships change only when one
 * of those additions declares itself through the context every planner shares.
 *
 * The body is written before that relationships part, because a link asks for its relationship as
 * it is written (`docx/hyperlink`), and both writers hand out ids through the one writer so that
 * they cannot pick the same one.
 */

import type { Node as PMNode } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import {
  ensureRootDeclarations,
  type RootDeclarations,
} from "../ooxml/partSplice";
import {
  encodeUtf8,
  parseXml,
  R_NS,
  W_NS,
  withXmlParser,
  type XmlParser,
} from "../ooxml/xml";
import { sameSource } from "../schema/sourceEquality";
import { commentsPlanner } from "./comments";
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
import { numberingPlanner } from "./numberingPlanner";
import { CONTENT_TYPES_PATH, contentTypeWriter } from "./packageParts";
import {
  assertPartsParse,
  type PartPlanContext,
  type PartPlanner,
  runPartPlanners,
} from "./partPlan";
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
  return serializeBlock(node, refs);
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

/** Checks bookmark pairing and, when links were written, their relationship namespace in scope. */
function assertMainPart(documentXml: string, wroteLinks: boolean): void {
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
    if (wroteLinks && element.localName === "hyperlink") {
      const id = element.getAttributeNode("r:id");
      if (id && id.namespaceURI !== R_NS) {
        throw new DocxExportError(
          "unsupported-content",
          "a hyperlink's r:id is shadowed by a different relationship namespace"
        );
      }
    }
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
 * What the main part has to declare once a link has gone out under an `r:id` put on its tag: the
 * attribute names a relationship, and it names nothing at all in a part that binds no `r`.
 */
const LINK_MARKUP: RootDeclarations = { namespaces: { r: R_NS } };

/** The parts written beside the body, in the order their parts go into the package */
const PART_PLANNERS: readonly PartPlanner[] = [
  numberingPlanner,
  commentsPlanner,
];

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
  return reportThrough(PART_PLANNERS, doc, session, options);
}

/**
 * The same export, folding a planner list of the caller's choosing in place of `PART_PLANNERS`.
 *
 * On no entry point. No document can make a correct planner write a part that does not read back,
 * so this is how a test hands the fold such a planner and watches the export refuse it.
 */
export function exportThroughPlanners(
  planners: readonly PartPlanner[],
  doc: PMNode,
  session: DocxSession,
  options?: ExportOptions
): Uint8Array {
  return reportThrough(planners, doc, session, options).bytes;
}

function reportThrough(
  planners: readonly PartPlanner[],
  doc: PMNode,
  session: DocxSession,
  options: ExportOptions | undefined
): { bytes: Uint8Array; notes: FidelityNote[] } {
  return withXmlParser(options?.xmlParser, () => {
    const store = sessionOf(session);
    const problem = problemsOf(doc, store)[0];
    if (problem) throw new DocxExportError(problem.code, problem.message);
    const approximated: FidelityNote[] = [];
    const bytes = writeDocx(
      doc,
      store,
      { add: (note) => approximated.push(note) },
      planners
    );
    return {
      bytes,
      notes: [...fidelityNotesOf(doc, store.mainPartPath), ...approximated],
    };
  });
}

function writeDocx(
  doc: PMNode,
  store: SessionStore,
  notes: FidelityCollector,
  planners: readonly PartPlanner[]
): Uint8Array {
  const relsPath = relsPathOf(store.mainPartPath);
  const context: PartPlanContext = {
    relationships: relationshipWriter(readRelationships(store.parts, relsPath)),
    contentTypes: contentTypeWriter(store.parts),
  };
  // The body has to know which relationship a newly inserted image ends up on, so the
  // media is planned before the body is written
  const media = planImageMedia(doc, store, context);
  const links = hyperlinkRefs(context.relationships);
  const body = buildDocumentXml(doc, store, {
    images: media?.refs ?? NO_IMAGE_REFS,
    links,
    notes,
    session: store,
  });
  const documentXml = links.addedRelId()
    ? ensureRootDeclarations(body, LINK_MARKUP)
    : body;
  assertMainPart(documentXml, links.addedRelId());

  const parts = runPartPlanners(planners, doc, store, context, media?.parts);
  const rels = context.relationships.part(store.parts.get(relsPath));
  if (rels) parts.set(relsPath, rels);
  const contentTypes = context.contentTypes.part();
  if (contentTypes) parts.set(CONTENT_TYPES_PATH, contentTypes);
  // The body was read back by assertBookmarkPairs; every other rewritten part is read back here
  assertPartsParse(parts, contentTypes ?? store.parts.get(CONTENT_TYPES_PATH));

  return repackParts(
    store.parts,
    new Map([
      [store.mainPartPath, encodeUtf8(documentXml, store.documentHadBom)],
      ...parts,
    ])
  );
}
