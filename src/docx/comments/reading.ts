/**
 * Reads comment and comment-extension package parts.
 */

import {
  attributeByLocalName,
  decodeUtf8,
  elementChildren,
  parseXml,
  serializeXml,
} from "../../ooxml/xml";
import { relatedPartPath } from "../packageParts";
import {
  COMMENTS_EXTENDED_REL_TYPE,
  COMMENTS_EXTENSIBLE_REL_TYPE,
  COMMENTS_IDS_REL_TYPE,
  COMMENTS_REL_TYPE,
} from "./constants";
import { lastBodyParagraph } from "./grammar";
import {
  commentAuthorId,
  type ImportedPeople,
  NO_PEOPLE,
  readPeople,
} from "./people";

/**
 * One entry of the Comments part as it arrived, everything about it but what it says: that is a
 * story of its own, sliced and read the way a body block is (`docx/story`).
 */
export interface ImportedComment {
  id: string;
  author: string | null;
  /**
   * The identity the people part records for `author` under this editor's provider. Null when it
   * records none, and null as well when it records that name under more than one identity, which
   * leaves no way to tell whose comment this is.
   */
  authorId: string | null;
  initials: string | null;
  /** As written: the author's wall clock (`./dates`) */
  date: string | null;
  paraId: string | null;
  parentParaId: string | null;
  resolved: boolean;
  extensionXml: string | null;
  durableId: string | null;
  /** As written, where the durable id reaches one */
  dateUtc: string | null;
}

/** A `w16cid:commentId`, keyed by the thread key of the comment it stands for */
export interface ImportedCommentId {
  paraId: string;
  durableId: string;
}

/** A `w16cex:commentExtensible` */
export interface ImportedCommentDate {
  durableId: string;
  dateUtc: string | null;
}

export interface ImportedCommentPart<Entry> {
  partPath: string | null;
  xml: string | null;
  hadBom: boolean;
  ordered: readonly Entry[];
}

const NO_PART: ImportedCommentPart<never> = {
  partPath: null,
  xml: null,
  hadBom: false,
  ordered: [],
};

export interface ImportedComments {
  partPath: string | null;
  xml: string | null;
  hadBom: boolean;
  ordered: readonly ImportedComment[];
  byId: ReadonlyMap<string, ImportedComment>;
  repliesByParentId: ReadonlyMap<string, readonly ImportedComment[]>;
  extendedPartPath: string | null;
  extendedXml: string | null;
  extendedHadBom: boolean;
  extendedOrdered: readonly ImportedCommentExtension[];
  /** The people part as opened, which is where an author's identity is looked up */
  people: ImportedPeople;
  ids: ImportedCommentPart<ImportedCommentId>;
  extensible: ImportedCommentPart<ImportedCommentDate>;
}

export const NO_COMMENTS: ImportedComments = {
  partPath: null,
  xml: null,
  hadBom: false,
  ordered: [],
  byId: new Map(),
  repliesByParentId: new Map(),
  extendedPartPath: null,
  extendedXml: null,
  extendedHadBom: false,
  extendedOrdered: [],
  people: NO_PEOPLE,
  ids: NO_PART,
  extensible: NO_PART,
};

export interface ImportedCommentExtension {
  paraId: string;
  parentParaId: string | null;
  resolved: boolean;
  xml: string;
}

/** The thread key of a comment, which the writer puts on the last paragraph of its body */
export function lastParagraphId(comment: Element): string | null {
  const last = lastBodyParagraph(comment);
  return last === null ? null : attributeByLocalName(last, "paraId");
}

function readCommentPart<Entry>(
  parts: Map<string, Uint8Array>,
  mainPartPath: string,
  relType: string,
  localName: string,
  entryOf: (el: Element) => Entry | null
): ImportedCommentPart<Entry> {
  const partPath = relatedPartPath(parts, mainPartPath, relType);
  if (partPath === null) return NO_PART;
  const bytes = parts.get(partPath);
  if (!bytes) return { ...NO_PART, partPath };
  const { text, hadBom } = decodeUtf8(bytes);
  const ordered = elementChildren(parseXml(text).documentElement).flatMap(
    (el) => {
      const entry = el.localName === localName ? entryOf(el) : null;
      return entry === null ? [] : [entry];
    }
  );
  return { partPath, xml: text, hadBom, ordered };
}

function commentExtensionOf(el: Element): ImportedCommentExtension | null {
  const paraId = attributeByLocalName(el, "paraId");
  return paraId === null
    ? null
    : {
        paraId,
        parentParaId: attributeByLocalName(el, "paraIdParent"),
        resolved: ["1", "true", "on"].includes(
          attributeByLocalName(el, "done")?.toLowerCase() ?? ""
        ),
        xml: serializeXml(el),
      };
}

function commentIdOf(el: Element): ImportedCommentId | null {
  const paraId = attributeByLocalName(el, "paraId");
  const durableId = attributeByLocalName(el, "durableId");
  return paraId === null || durableId === null ? null : { paraId, durableId };
}

function commentDateOf(el: Element): ImportedCommentDate | null {
  const durableId = attributeByLocalName(el, "durableId");
  return durableId === null
    ? null
    : { durableId, dateUtc: attributeByLocalName(el, "dateUtc") };
}

/** The first entry under a key wins, as in every other comment part */
function firstByKey<Entry, Value>(
  entries: readonly Entry[],
  key: (entry: Entry) => string,
  value: (entry: Entry) => Value
): ReadonlyMap<string, Value> {
  const found = new Map<string, Value>();
  for (const entry of entries) {
    if (!found.has(key(entry))) found.set(key(entry), value(entry));
  }
  return found;
}

/** A durable id is hexadecimal (`ST_LongHexNumber`), so two spellings in different case are one id */
export function durableIdKey(durableId: string): string {
  return durableId.toUpperCase();
}

/** Every durable id the ids and extensible parts name, orphans included: a date left under one would become a new comment's */
export function spentDurableIds(
  comments: ImportedComments
): ReadonlySet<string> {
  return new Set(
    [...comments.ids.ordered, ...comments.extensible.ordered].map((entry) =>
      durableIdKey(entry.durableId)
    )
  );
}

/**
 * Reads the Comments part related from the main document story.
 *
 * The people part and the extended part are read whether or not there are comments: a document may
 * carry either with no comment left, and a comment added to it then has to be written into the
 * part it already has rather than into a second one beside it.
 */
export function readComments(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): ImportedComments {
  const people = readPeople(parts, mainPartPath);
  const extensions = readCommentPart(
    parts,
    mainPartPath,
    COMMENTS_EXTENDED_REL_TYPE,
    "commentEx",
    commentExtensionOf
  );
  const extensionsByParaId = firstByKey(
    extensions.ordered,
    (entry) => entry.paraId,
    (entry) => entry
  );
  const ids = readCommentPart(
    parts,
    mainPartPath,
    COMMENTS_IDS_REL_TYPE,
    "commentId",
    commentIdOf
  );
  const extensible = readCommentPart(
    parts,
    mainPartPath,
    COMMENTS_EXTENSIBLE_REL_TYPE,
    "commentExtensible",
    commentDateOf
  );
  const aside = {
    people,
    extendedPartPath: extensions.partPath,
    extendedXml: extensions.xml,
    extendedHadBom: extensions.hadBom,
    extendedOrdered: extensions.ordered,
    ids,
    extensible,
  };
  const durableIds = firstByKey(
    ids.ordered,
    (entry) => entry.paraId,
    (entry) => entry.durableId
  );
  const utcDates = firstByKey(
    extensible.ordered,
    (entry) => durableIdKey(entry.durableId),
    (entry) => entry.dateUtc
  );
  const partPath = relatedPartPath(parts, mainPartPath, COMMENTS_REL_TYPE);
  if (partPath === null) return { ...NO_COMMENTS, ...aside };

  const bytes = parts.get(partPath);
  if (!bytes) return { ...NO_COMMENTS, partPath, ...aside };

  const { text, hadBom } = decodeUtf8(bytes);
  const root = parseXml(text).documentElement;
  const base = elementChildren(root)
    .filter((el) => el.localName === "comment")
    .flatMap((el) => {
      const id = attributeByLocalName(el, "id");
      if (id === null) return [];
      const paraId = lastParagraphId(el);
      const extension = paraId ? extensionsByParaId.get(paraId) : undefined;
      const author = attributeByLocalName(el, "author");
      const durableId =
        paraId === null ? null : (durableIds.get(paraId) ?? null);
      return [
        {
          id,
          author,
          authorId: author === null ? null : commentAuthorId(people, author),
          initials: attributeByLocalName(el, "initials"),
          date: attributeByLocalName(el, "date"),
          paraId,
          parentParaId: extension?.parentParaId ?? null,
          resolved: extension?.resolved ?? false,
          extensionXml: extension?.xml ?? null,
          durableId,
          dateUtc:
            durableId === null
              ? null
              : (utcDates.get(durableIdKey(durableId)) ?? null),
        },
      ];
    });
  const idByParaId = new Map(
    base.flatMap((comment) =>
      comment.paraId === null ? [] : [[comment.paraId, comment.id] as const]
    )
  );
  const ordered = base;
  const byId = new Map<string, ImportedComment>();
  const repliesByParentId = new Map<string, ImportedComment[]>();
  for (const comment of ordered) {
    if (!byId.has(comment.id)) byId.set(comment.id, comment);
    const parentId =
      comment.parentParaId === null
        ? null
        : (idByParaId.get(comment.parentParaId) ?? null);
    if (parentId !== null) {
      const replies = repliesByParentId.get(parentId) ?? [];
      replies.push(comment);
      repliesByParentId.set(parentId, replies);
    }
  }
  return {
    partPath,
    xml: text,
    hadBom,
    ordered,
    byId,
    repliesByParentId,
    ...aside,
  };
}
