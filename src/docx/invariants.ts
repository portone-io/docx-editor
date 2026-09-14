/**
 * Known export refusals, asked of the document before anything is written.
 *
 * Each invariant reads what the writer would refuse over and reports it with the code the writer
 * would throw, so a screen can ask at edit time what `exportDocx` would say. The list is walked in
 * order and `exportDocx` throws its first entry, which is what keeps the two from disagreeing.
 *
 * A code covers several situations, so each entry also carries the `ExportProblemReason`
 * (`ooxml/errors`) naming the one that was met and the content it is about. Every place a problem
 * is built names its own, rather than a reader telling them apart by the message.
 *
 * The checks inspect model attrs and preserved XML without running the writers. Bookmark
 * fragments and list definitions use the export parser. `assertBookmarkPairs` in `./exportDocx`
 * and `assertPartsParse` in `./partPlan` remain the final checks over the parts as actually written.
 */

import type { Node as PMNode } from "prosemirror-model";
import { spanCount, toParagraphFormat } from "../model/format";
import type {
  ExportProblem,
  ExportProblemReason,
  ExportProblemStory,
} from "../ooxml/errors";
import {
  attributeByLocalName,
  parseXml,
  W_NS,
  withXmlParser,
} from "../ooxml/xml";
import { visitPreservedFragments } from "../schema/preservedFragments";
import { unattributedCommentAuthors } from "../schema/protection";
import {
  HEADER_FOOTER_KINDS,
  NOTE_KINDS,
  type NoteKind,
  type StoryKey,
  type StoryKind,
  storyKey,
} from "../schema/stories";

import {
  commentReferencesIn,
  commentsChanged,
  commentsRootProblem,
  extensionsChanged,
  extensionsRootProblem,
} from "./comments";
import {
  commentsExtendedPart,
  commentsPart,
  peoplePart,
} from "./comments/parts";
import { unrecordedAuthors } from "./comments/people";
import { currentCommentBodies } from "./comments/writing";
import type { ExportOptions } from "./exportDocx";
import { identityProblems, identityProblemsInStories } from "./identities";
import { insertedImageSrcs } from "./media";
import { canDefineNewList, newNumIds, startedLists } from "./newLists";
import { CONTENT_TYPES_PATH } from "./packageParts";
import { STORY_ENTRIES_PARTS, STORY_WRITINGS } from "./partPlanners";
import { lostOriginal } from "./serializePreserved";
import {
  type DocxSession,
  originalBlock,
  type SessionStore,
  sessionOf,
} from "./session";
import {
  type StoryEntry,
  storyChangesOf,
  storyEntriesOf,
  storyEntriesProblems,
  storyIdOf,
  unwrittenStoryChanges,
} from "./storyParts";

export type {
  ExportPartName,
  ExportProblem,
  ExportProblemReason,
  ExportProblemStory,
  ExportStoryKind,
} from "../ooxml/errors";

interface ExportInvariant {
  readonly name: string;
  check(doc: PMNode, session: SessionStore): readonly ExportProblem[];
}

/** OOXML requires every bookmark end to identify an earlier unmatched start, and every start to be ended */
const bookmarkPairs: ExportInvariant = {
  name: "bookmarkPairs",
  check(doc, session) {
    const problems: ExportProblem[] = [];
    const open = new Map<string, number>();
    const used = new Set<string>();
    visitPreservedFragments(doc, (node, pos, xml) => {
      const source = xml ?? originalBlock(node, session)?.xml ?? null;
      if (source === null) return;
      // Use the document's namespace scope, and let the parser distinguish elements from
      // comments/CDATA and decode attribute references just as the final writer check does.
      if (!source.includes("bookmark")) return;
      let root: Element;
      try {
        root = parseXml(
          session.documentPrefix + source + session.documentSuffix
        ).documentElement;
      } catch {
        problems.push({
          code: "malformed-xml",
          message: "preserved bookmark XML could not be parsed",
          reason: { kind: "unreadable-preserved-xml" },
          pos,
        });
        return;
      }
      for (const marker of Array.from(root.getElementsByTagName("*"))) {
        if (
          marker.namespaceURI !== W_NS ||
          (marker.localName !== "bookmarkStart" &&
            marker.localName !== "bookmarkEnd")
        )
          continue;
        const kind = marker.localName.slice("bookmark".length);
        const value = attributeByLocalName(marker, "id");
        if (value === null || value === "") {
          problems.push({
            code: "malformed-xml",
            message: `a bookmark${kind} has no id`,
            reason: {
              kind: "unnamed-bookmark",
              marker: kind === "Start" ? "start" : "end",
            },
            pos,
          });
        } else if (kind === "Start") {
          if (used.has(value)) {
            problems.push({
              code: "malformed-xml",
              message: `bookmark ${value} has more than one start marker`,
              reason: { kind: "repeated-bookmark-start", id: value },
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
            reason: {
              kind: "unmatched-bookmark",
              id: value,
              marker: "end",
            },
            pos,
          });
        }
      }
    });
    for (const [id, pos] of open) {
      problems.push({
        code: "malformed-xml",
        message: `bookmark ${id} has no end marker`,
        reason: { kind: "unmatched-bookmark", id, marker: "start" },
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
          reason: { kind: "vertical-merge-past-table" },
          pos,
        });
      }
      return true;
    });
    return problems;
  },
};

/** The preserved nodes that always carry their own fragment rather than pointing at an original block */
const CARRIES_ITS_XML: ReadonlySet<string> = new Set([
  "rawInline",
  "rawRunContent",
]);

/** Why a preserved node has nothing to be written from, or null when it has */
function lostOriginalOf(
  node: PMNode,
  session: SessionStore
): { readonly message: string; readonly reason: ExportProblemReason } | null {
  if (CARRIES_ITS_XML.has(node.type.name)) {
    return typeof node.attrs.xml === "string"
      ? null
      : {
          message: "a preserved element has lost its original XML",
          reason: { kind: "lost-preserved-xml", node: node.type.name },
        };
  }
  if (node.type.isInGroup("preserved")) {
    if (typeof node.attrs.xml === "string") return null;
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
      const lost = lostOriginalOf(node, session);
      if (lost !== null) {
        problems.push({ code: "lost-original", ...lost, pos });
      }
      return true;
    });
    return problems;
  },
};

/** The stories one part settles its names across, all of them of the one kind the part holds */
interface StoryPartEntries {
  readonly kind: StoryKind;
  readonly entries: readonly StoryEntry[];
}

/** The story a key names, and null where there is no key to name one */
function storyNameOf(
  key: StoryKey | null,
  kind: StoryKind
): ExportProblemStory | null {
  return key === null ? null : { kind, id: storyIdOf(key, kind) };
}

/**
 * A name held by one node only is settled by `withUniqueIdentities` just before the body is
 * written, and by `withUniqueStoryIdentities` before a header, footer, or footnotes part an edit
 * changed is written again: a later claimant is rebuilt from its own attrs, and a block preserved
 * as nothing but its original XML has nothing to be rebuilt from, so the pass refuses it. The pass
 * is asked here rather than read again, so the block it names is the one the write would refuse
 * over. A block of a side story stands nowhere in the body, so its problem is reported where the
 * body refers to the note it belongs to, and nowhere at all for a header or a footer.
 */
const uniqueIdentities: ExportInvariant = {
  name: "uniqueIdentities",
  check(doc, session) {
    const parts: readonly StoryPartEntries[] = [
      ...HEADER_FOOTER_KINDS.flatMap((kind) =>
        storyChangesOf(doc, session, kind).flatMap(
          (change): StoryPartEntries[] =>
            change.change === "edited"
              ? [
                  {
                    kind,
                    entries: [
                      {
                        key: change.imported.key,
                        story: change.current,
                        frozen: false,
                      },
                    ],
                  },
                ]
              : []
        )
      ),
      ...STORY_ENTRIES_PARTS.map((part) => ({
        kind: part.kind,
        entries: storyEntriesOf(part, doc, session),
      })),
    ];
    return [
      ...identityProblems(doc).map(
        ({ code, message, node, pos }): ExportProblem => ({
          code,
          message,
          reason: { kind: "duplicate-preserved-block", node, story: null },
          pos,
        })
      ),
      ...parts.flatMap(({ kind, entries }) =>
        identityProblemsInStories(entries).map(
          ({ code, message, node, story }): ExportProblem => ({
            code,
            message,
            reason: {
              kind: "duplicate-preserved-block",
              node,
              story: storyNameOf(entries[story]?.key ?? null, kind),
            },
          })
        )
      ),
    ];
  },
};

/**
 * A story changes on the document node whether or not a part writer carries it into the file. A
 * change nothing writes - to an endnote, to a header story added or removed, or to a separator
 * entry, which goes back out as it arrived - is refused here rather than dropped from the file
 * without a word.
 */
const storiesHaveWriters: ExportInvariant = {
  name: "storiesHaveWriters",
  check(doc, session) {
    return unwrittenStoryChanges(STORY_WRITINGS, doc, session).map(
      ({ key, kind, id, change }): ExportProblem => ({
        code: "unsupported-content",
        message: `the ${key} story was ${change}, and no part writer carries that into the file`,
        reason: {
          kind: "unwritten-story-change",
          story: { kind, id },
          change,
        },
      })
    );
  },
};

/**
 * A list started while editing goes out as the definition it was registered with
 * (`numbering/listRegistry`). A paragraph in a list nothing defines - not the file it was opened
 * from, and not the register the document node carries - has no definition to be written, and the
 * file would go out naming a list defined nowhere, which is the one failure the register exists to
 * rule out.
 *
 * The code is the one that already covers a document holding what no correct file can be written
 * from: nothing is missing from the package, and the paragraph has lost nothing of its own. Only
 * markup written by something other than the list commands reaches it, the same caller the code's
 * other case names.
 */
const listDefinitions: ExportInvariant = {
  name: "listDefinitions",
  check(doc, session) {
    const undefinedIds = new Set(startedLists(doc, session).unregistered);
    if (undefinedIds.size === 0) return [];
    const problems: ExportProblem[] = [];
    doc.descendants((node, pos) => {
      if (node.type.name !== "paragraph") return true;
      const numId = toParagraphFormat(node.attrs.format)?.numbering?.numId;
      if (numId === undefined || !undefinedIds.delete(numId)) return false;
      problems.push({
        code: "unsupported-content",
        message: `the list numbered ${numId} has no definition to be written`,
        reason: { kind: "undefined-list", numId },
        pos,
      });
      return false;
    });
    return problems;
  },
};

/** Whether a changed comment needs a part the package has yet to declare. */
function addsCommentsPart(doc: PMNode, session: SessionStore): boolean {
  const bodyChanged = commentsChanged(doc, session);
  const threadChanged = extensionsChanged(doc, session);
  if (!bodyChanged && !threadChanged) return false;
  if (commentsPart.pathIn(session) === null || session.comments.xml === null)
    return true;
  const references = commentReferencesIn(doc);
  if (
    threadChanged &&
    (references.size > 0 || session.comments.extendedPartPath !== null) &&
    (commentsExtendedPart.pathIn(session) === null ||
      session.comments.extendedXml === null)
  )
    return true;
  return (
    bodyChanged &&
    (peoplePart.pathIn(session) === null ||
      session.comments.people.xml === null) &&
    unrecordedAuthors(
      currentCommentBodies(references).values(),
      session.comments.people,
      unattributedCommentAuthors(session.comments.ordered)
    ).size > 0
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
        reason: { kind: "missing-content-types", part: "media" },
      });
    }
    if (!canDefineNewList(session) && newNumIds(doc, session).length > 0) {
      problems.push({
        code: "missing-content-types",
        message: `cannot add a part to a package that has no ${CONTENT_TYPES_PATH}`,
        reason: { kind: "missing-content-types", part: "numbering" },
      });
    }
    if (addsCommentsPart(doc, session)) {
      problems.push({
        code: "missing-content-types",
        message: `cannot add a part to a package that has no ${CONTENT_TYPES_PATH}`,
        reason: { kind: "missing-content-types", part: "comments" },
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
      if (problem !== null) {
        problems.push({
          code: "malformed-xml",
          message: problem,
          reason: { kind: "unwritable-part-root", part: "comments" },
        });
      }
    }
    if (
      extendedXml !== null &&
      extensionsChanged(doc, session) &&
      (commentReferencesIn(doc).size > 0 || extendedPartPath !== null)
    ) {
      const problem = extensionsRootProblem(extendedXml);
      if (problem !== null) {
        problems.push({
          code: "malformed-xml",
          message: problem,
          reason: { kind: "unwritable-part-root", part: "commentsExtended" },
        });
      }
    }
    return problems;
  },
};

/**
 * A part written one entry per story is rewritten around its root element, so a part with none
 * cannot take a changed story, and a part the package lacks is declared in the content types part.
 * An untouched part is never read, which is how a document whose notes nobody changed exports
 * whatever its notes parts look like.
 */
const notePartRoots: ExportInvariant = {
  name: "notePartRoots",
  check: (doc, session) =>
    STORY_ENTRIES_PARTS.flatMap((part) =>
      storyEntriesProblems(part, doc, session)
    ),
};

/**
 * Where the body first refers to each note, under the key naming that note.
 *
 * A note's text is a story of its own and stands nowhere in the body, so a problem about one is
 * reported at the reference a reader would have to look at to decide what to do about it.
 */
function noteReferencePositions(doc: PMNode): ReadonlyMap<StoryKey, number> {
  const first = new Map<StoryKey, number>();
  doc.descendants((node, pos) => {
    if (node.type.name !== "noteReference") return true;
    const kind = NOTE_KINDS.find((candidate) => candidate === node.attrs.kind);
    const id: unknown = node.attrs.id;
    if (kind !== undefined && typeof id === "string") {
      const key = storyKey(kind, id);
      if (!first.has(key)) first.set(key, pos);
    }
    return true;
  });
  return first;
}

/** The story a reason is about, and null for one about the body or the package */
function storyOf(reason: ExportProblemReason): ExportProblemStory | null {
  switch (reason.kind) {
    case "duplicate-preserved-block":
    case "story-id-not-a-number":
    case "unwritten-story-change":
      return reason.story;
    default:
      return null;
  }
}

/** The note a reason is about, and null for one about no note */
function noteOf(
  reason: ExportProblemReason
): { readonly kind: NoteKind; readonly id: string } | null {
  const story = storyOf(reason);
  if (story === null) return null;
  const kind = NOTE_KINDS.find((candidate) => candidate === story.kind);
  return kind === undefined ? null : { kind, id: story.id };
}

/** Every problem about a note reported where the body refers to that note, the rest as they stand */
function atNoteReferences(
  problems: readonly ExportProblem[],
  doc: PMNode
): readonly ExportProblem[] {
  const notes = problems.map((problem) =>
    problem.pos === undefined ? noteOf(problem.reason) : null
  );
  if (notes.every((note) => note === null)) return problems;
  const references = noteReferencePositions(doc);
  return problems.map((problem, at) => {
    const note = notes[at];
    const pos = note ? references.get(storyKey(note.kind, note.id)) : undefined;
    return pos === undefined ? problem : { ...problem, pos };
  });
}

/** In the order the problems are reported, which is the order `exportDocx` throws them in */
const EXPORT_INVARIANTS: readonly ExportInvariant[] = [
  bookmarkPairs,
  tableGrids,
  preservedOriginals,
  uniqueIdentities,
  listDefinitions,
  storiesHaveWriters,
  mediaContentTypes,
  commentPartRoots,
  notePartRoots,
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
  const problems = atNoteReferences(
    EXPORT_INVARIANTS.flatMap((invariant) => invariant.check(doc, session)),
    doc
  );
  answered.set(doc, { session, problems });
  return problems;
}

/**
 * Known reasons writing this document back would be refused, in the order `exportDocx` would
 * raise them. An empty list does not rule out failures while writing. Each entry carries the code
 * and message the `DocxExportError` would carry, the `reason` naming what the refusal is about,
 * and the position in the document where there is one.
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
