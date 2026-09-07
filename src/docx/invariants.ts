/**
 * Every reason an export would be refused, asked of the document before anything is written.
 *
 * Each invariant reads what the writer would refuse over and reports it with the code the writer
 * would throw, so a screen can ask at edit time what `exportDocx` would say. The list is walked in
 * order and `exportDocx` throws its first entry, which is what keeps the two from disagreeing.
 *
 * The invariants read nothing but node attrs and the original XML the session holds, as strings:
 * no part is parsed and nothing is assembled. `assertBookmarkPairs` in `./exportDocx` stays as the
 * last line of defence over the body as it was actually written.
 */

import type { Node as PMNode } from "prosemirror-model";
import { spanCount } from "../model/format";
import type { DocxExportErrorCode } from "../ooxml/errors";
import { withXmlParser } from "../ooxml/xml";
import {
  commentReferencesIn,
  commentsChanged,
  commentsRootProblem,
  extensionsChanged,
  extensionsRootProblem,
} from "./comments";
import { CONTENT_TYPES_PATH } from "./comments/constants";
import { commentsPart } from "./comments/parts";
import type { ExportOptions } from "./exportDocx";
import { insertedImageSrcs } from "./media";
import { newNumIds, numberingPartOf } from "./newLists";
import { lostOriginal } from "./serializeBlock";
import {
  type DocxSession,
  originalBlock,
  type SessionStore,
  sessionOf,
} from "./session";

/** One reason the document cannot be written back, with the code `exportDocx` would throw it under */
export interface ExportProblem {
  readonly code: DocxExportErrorCode;
  readonly message: string;
  /** Where the problem stands in the document. Absent for a problem of the package or of the session */
  readonly pos?: number;
}

interface ExportInvariant {
  readonly name: string;
  check(doc: PMNode, session: SessionStore): readonly ExportProblem[];
}

const BOOKMARK_MARKER = /<(?:[\w.-]+:)?bookmark(Start|End)\b([^>]*)>/g;
const ID_ATTRIBUTE = /\s(?:[\w.-]+:)?id\s*=\s*(?:"([^"]*)"|'([^']*)')/;

/**
 * The XML a node's bookmark markers stand in, or null for a node that holds none.
 *
 * A marker inside a paragraph is a `rawInline` of its own, one directly under the body is a
 * `bookmarkBlock` pointing at its original, and one inside an unsupported container stays in that
 * container's XML (spec/notes/bookmarks.md "What we preserve"), which a `docxRaw` points at and a
 * `rawBlock` carries.
 */
function markerSourceOf(node: PMNode, session: SessionStore): string | null {
  if (node.type.name === "rawInline" || node.type.name === "rawBlock") {
    return typeof node.attrs.xml === "string" ? node.attrs.xml : null;
  }
  if (node.type.isInGroup("preserved")) {
    return originalBlock(node, session)?.xml ?? null;
  }
  return null;
}

/** OOXML requires every bookmark end to identify an earlier unmatched start, and every start to be ended */
const bookmarkPairs: ExportInvariant = {
  name: "bookmarkPairs",
  check(doc, session) {
    const problems: ExportProblem[] = [];
    const open = new Map<string, number>();
    const used = new Set<string>();
    doc.descendants((node, pos) => {
      const source = markerSourceOf(node, session);
      if (source === null) return true;
      for (const marker of source.matchAll(BOOKMARK_MARKER)) {
        const kind = marker[1];
        const id = ID_ATTRIBUTE.exec(marker[2] ?? "");
        const value = id?.[1] ?? id?.[2];
        if (value === undefined) {
          problems.push({
            code: "malformed-xml",
            message: `a bookmark${kind} has no id`,
            pos,
          });
        } else if (kind === "Start") {
          if (used.has(value)) {
            problems.push({
              code: "malformed-xml",
              message: `bookmark ${value} has more than one start marker`,
              pos,
            });
          } else {
            open.set(value, pos);
            used.add(value);
          }
        } else if (!open.delete(value)) {
          problems.push({
            code: "malformed-xml",
            message: `bookmark ${value} ends without an earlier start marker`,
            pos,
          });
        }
      }
      return true;
    });
    for (const [id, pos] of open) {
      problems.push({
        code: "malformed-xml",
        message: `bookmark ${id} has no end marker`,
        pos,
      });
    }
    return problems;
  },
};

/** Whether a cell's vertical merge covers rows the table does not have */
function mergeReachesPastLastRow(table: PMNode): boolean {
  let overlong = false;
  table.forEach((row, _offset, at) => {
    row.forEach((cell) => {
      if (at + spanCount(cell.attrs.rowspan) > table.childCount) {
        overlong = true;
      }
    });
  });
  return overlong;
}

/**
 * A vertical merge reaching past the last row is the one grid the table writer refuses: cells
 * that overlap or a row that comes up short is a table it writes as it stands. The core entry
 * reaches no table library, so the rows are counted here rather than read off a `TableMap`.
 */
const tableGrids: ExportInvariant = {
  name: "tableGrids",
  check(doc) {
    const problems: ExportProblem[] = [];
    doc.descendants((node, pos) => {
      if (node.type.name !== "table") return true;
      if (mergeReachesPastLastRow(node)) {
        problems.push({
          code: "invalid-table",
          message: "a vertical merge in the table reaches past the last row",
          pos,
        });
      }
      return true;
    });
    return problems;
  },
};

/** Why a preserved node has nothing to be written from, or null when it has */
function lostOriginalOf(node: PMNode, session: SessionStore): string | null {
  if (node.type.name === "rawInline" || node.type.name === "rawBlock") {
    return typeof node.attrs.xml === "string"
      ? null
      : "a preserved element has lost its original XML";
  }
  if (node.type.isInGroup("preserved")) {
    return originalBlock(node, session) ? null : lostOriginal(node, session);
  }
  return null;
}

/** A preserved node is written from its original XML, which it either carries or points at in this session */
const preservedOriginals: ExportInvariant = {
  name: "preservedOriginals",
  check(doc, session) {
    const problems: ExportProblem[] = [];
    doc.descendants((node, pos) => {
      const message = lostOriginalOf(node, session);
      if (message !== null)
        problems.push({ code: "lost-original", message, pos });
      return true;
    });
    return problems;
  },
};

/**
 * A new list is defined by splicing into numbering.xml. Creating that part from scratch would
 * also mean adding a part and touching up [Content_Types].xml, so a document without one cannot
 * take a new list, and this says so rather than the export handing back a half-finished file.
 */
const numberingPart: ExportInvariant = {
  name: "numberingPart",
  check(doc, session) {
    if (numberingPartOf(session) !== null) return [];
    if (newNumIds(doc, session).length === 0) return [];
    return [
      {
        code: "missing-numbering-part",
        message:
          "cannot add a new list to a document that has no numbering.xml",
      },
    ];
  },
};

/** Whether the export would write a Comments part the package has yet to declare */
function addsCommentsPart(doc: PMNode, session: SessionStore): boolean {
  return (
    (commentsChanged(doc, session) || extensionsChanged(doc, session)) &&
    (commentsPart.pathIn(session) === null || session.comments.xml === null)
  );
}

/**
 * A part the export adds has to be declared in [Content_Types].xml, and writing that part from
 * scratch would mean guessing the type of every other part in the package.
 */
const mediaContentTypes: ExportInvariant = {
  name: "mediaContentTypes",
  check(doc, session) {
    if (session.parts.has(CONTENT_TYPES_PATH)) return [];
    const problems: ExportProblem[] = [];
    if (insertedImageSrcs(doc).length > 0) {
      problems.push({
        code: "missing-content-types",
        message: `cannot add an image to a package that has no ${CONTENT_TYPES_PATH}`,
      });
    }
    if (addsCommentsPart(doc, session)) {
      problems.push({
        code: "missing-content-types",
        message: `cannot add a part to a package that has no ${CONTENT_TYPES_PATH}`,
      });
    }
    return problems;
  },
};

/**
 * A comment part is rewritten around its root element, so a part with none cannot take a change.
 * An untouched part is never read, which is how a document nobody commented on exports whatever
 * its comment parts look like.
 */
const commentPartRoots: ExportInvariant = {
  name: "commentPartRoots",
  check(doc, session) {
    const problems: ExportProblem[] = [];
    const { xml, extendedXml, extendedPartPath } = session.comments;
    if (xml !== null && commentsChanged(doc, session)) {
      const problem = commentsRootProblem(xml);
      if (problem !== null)
        problems.push({ code: "malformed-xml", message: problem });
    }
    if (
      extendedXml !== null &&
      extensionsChanged(doc, session) &&
      (commentReferencesIn(doc).size > 0 || extendedPartPath !== null)
    ) {
      const problem = extensionsRootProblem(extendedXml);
      if (problem !== null)
        problems.push({ code: "malformed-xml", message: problem });
    }
    return problems;
  },
};

/** In the order the problems are reported, which is the order `exportDocx` throws them in */
const EXPORT_INVARIANTS: readonly ExportInvariant[] = [
  bookmarkPairs,
  tableGrids,
  preservedOriginals,
  numberingPart,
  mediaContentTypes,
  commentPartRoots,
];

/**
 * The answers already given, per document node. A document node is immutable, so a state whose
 * selection alone moved asks about the same node and gets the same list back without a second
 * walk; a session other than the one it was answered for is asked afresh.
 */
const answered = new WeakMap<
  PMNode,
  { session: SessionStore; problems: readonly ExportProblem[] }
>();

/** The problems as the writer sees them, for a caller already inside the writer's parser scope */
export function problemsOf(
  doc: PMNode,
  session: SessionStore
): readonly ExportProblem[] {
  const known = answered.get(doc);
  if (known && known.session === session) return known.problems;
  const problems = EXPORT_INVARIANTS.flatMap((invariant) =>
    invariant.check(doc, session)
  );
  answered.set(doc, { session, problems });
  return problems;
}

/**
 * Every reason writing this document back would be refused, in the order `exportDocx` would
 * raise them; empty when it would go through. Each entry carries the code and message the
 * `DocxExportError` would carry, and the position in the document where there is one.
 *
 * The list definitions are read to tell a new list from one the document already had, so this
 * needs an XML parser the way `exportDocx` does and takes the same option.
 */
export function exportProblems(
  doc: PMNode,
  session: DocxSession,
  options?: ExportOptions
): readonly ExportProblem[] {
  return withXmlParser(options?.xmlParser, () =>
    problemsOf(doc, sessionOf(session))
  );
}
