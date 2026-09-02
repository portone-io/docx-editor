/**
 * Whether a file handed back changed in nothing but comments one author is allowed to have made.
 *
 * The editor's own refusal under a comment protection (`schema/protection`) is a courtesy to the
 * user: the browser holds the file, so the rule is held again here, where a server takes it in.
 * The judgement is over the package rather than over the document alone, since a submission is
 * free to rewrite anything the story does not carry - the paper the document is written on, its
 * styles, its headers - and a document comparison would see none of it.
 */

import type { Node as PMNode } from "prosemirror-model";
import { decodeUtf8, elementChildren, parseXml } from "../ooxml/xml";
import {
  changesOnlyComments,
  commentAdditionsBy,
  commentEditsOwned,
  commentIdentitiesKept,
  type EditableComments,
} from "../schema/protection";
import {
  COMMENTS_CONTENT_TYPE,
  COMMENTS_EXTENDED_CONTENT_TYPE,
  COMMENTS_EXTENDED_REL_TYPE,
  COMMENTS_REL_TYPE,
  CONTENT_TYPES_PATH,
  PEOPLE_CONTENT_TYPE,
  PEOPLE_REL_TYPE,
} from "./comments/constants";
import { type DocxBytes, importDocx } from "./importDocx";
import {
  type Relationship,
  readRelationships,
  relsPathOf,
  resolveTarget,
} from "./relationships";
import type { SessionStore } from "./session";

/**
 * Why a file is not the one it claims to be. `part-changed` and `relationship-changed` name the
 * part they were reached over; the other three are about the document story itself.
 */
export type CommentOnlyVerdict =
  | { ok: true }
  | {
      ok: false;
      reason: "body-changed" | "comment-not-owned" | "comment-author-forged";
    }
  | {
      ok: false;
      reason: "part-changed" | "relationship-changed";
      part: string;
    };

/** The three parts a comment is written across, which are the parts a comment edit may rewrite */
const COMMENT_REL_TYPES: readonly string[] = [
  COMMENTS_REL_TYPE,
  COMMENTS_EXTENDED_REL_TYPE,
  PEOPLE_REL_TYPE,
];

const COMMENT_CONTENT_TYPES: readonly string[] = [
  COMMENTS_CONTENT_TYPE,
  COMMENTS_EXTENDED_CONTENT_TYPE,
  PEOPLE_CONTENT_TYPE,
];

const refused = (
  reason: "body-changed" | "comment-not-owned" | "comment-author-forged"
): CommentOnlyVerdict => ({ ok: false, reason });

function sameBytes(before: Uint8Array, after: Uint8Array): boolean {
  return (
    before.length === after.length &&
    before.every((byte, index) => byte === after[index])
  );
}

/** Where the comment parts sit in this package, whether or not the file carries them yet */
function commentPartPaths(session: SessionStore): Set<string> {
  const related = readRelationships(
    session.parts,
    relsPathOf(session.mainPartPath)
  ).filter(
    (entry) => !entry.external && COMMENT_REL_TYPES.includes(entry.type)
  );
  return new Set(
    related.map((entry) => resolveTarget(session.mainPartPath, entry.target))
  );
}

/**
 * The main document part with the blocks the document model carries taken out of it.
 *
 * What is left is everything the document comparison cannot see: the namespaces the story is
 * written under, the section properties that set the paper and its margins, and every block kept
 * as the XML it arrived as. A comment is written inside a paragraph, so nothing a comment edit
 * writes reaches this text.
 */
function aroundTheStory(session: SessionStore): string {
  const preserved = session.blocks
    .filter(
      (block) =>
        block.node.type.name !== "paragraph" && block.node.type.name !== "table"
    )
    .map((block) => block.xml)
    .join("");
  return session.documentPrefix + preserved + session.documentSuffix;
}

/**
 * Whether the relationships of the main document part are the ones it arrived with, save for the
 * comment parts it may have gained. An id already handed out keeps pointing where it pointed.
 */
function relationshipsKept(
  before: readonly Relationship[],
  after: readonly Relationship[]
): boolean {
  const now = new Map(after.map((entry) => [entry.id, entry]));
  const kept = before.every((entry) => {
    const current = now.get(entry.id);
    return (
      current !== undefined &&
      current.type === entry.type &&
      current.target === entry.target &&
      current.external === entry.external
    );
  });
  const ids = new Set(before.map((entry) => entry.id));
  return (
    kept &&
    after.every(
      (entry) => ids.has(entry.id) || COMMENT_REL_TYPES.includes(entry.type)
    )
  );
}

/** What `[Content_Types].xml` declares, each declaration under the name it is keyed by */
function contentTypes(bytes: Uint8Array | undefined): Map<string, string> {
  if (bytes === undefined) return new Map();
  const declared = new Map<string, string>();
  for (const el of elementChildren(
    parseXml(decodeUtf8(bytes).text).documentElement
  )) {
    const key =
      el.localName === "Default"
        ? el.getAttribute("Extension")
        : el.localName === "Override"
          ? el.getAttribute("PartName")
          : null;
    if (key !== null) declared.set(key, el.getAttribute("ContentType") ?? "");
  }
  return declared;
}

/**
 * Whether the package declares the content types it arrived with, save for an override a comment
 * part it gained needs. A declaration that was there keeps naming the type it named.
 */
function contentTypesKept(
  before: Map<string, string>,
  after: Map<string, string>
): boolean {
  for (const [key, type] of before) {
    if (after.get(key) !== type) return false;
  }
  for (const [key, type] of after) {
    if (!before.has(key) && !COMMENT_CONTENT_TYPES.includes(type)) return false;
  }
  return true;
}

/** Whether every part outside the document story is the one the file arrived with */
function packageKept(
  before: SessionStore,
  after: SessionStore
): CommentOnlyVerdict {
  if (before.mainPartPath !== after.mainPartPath) {
    return { ok: false, reason: "part-changed", part: before.mainPartPath };
  }
  const relsPath = relsPathOf(before.mainPartPath);
  const untouched = new Set([
    before.mainPartPath,
    relsPath,
    CONTENT_TYPES_PATH,
    ...commentPartPaths(before),
    ...commentPartPaths(after),
  ]);
  for (const path of new Set([...before.parts.keys(), ...after.parts.keys()])) {
    if (untouched.has(path)) continue;
    const was = before.parts.get(path);
    const now = after.parts.get(path);
    if (was === undefined || now === undefined || !sameBytes(was, now)) {
      return { ok: false, reason: "part-changed", part: path };
    }
  }
  if (
    !relationshipsKept(
      readRelationships(before.parts, relsPath),
      readRelationships(after.parts, relsPath)
    )
  ) {
    return { ok: false, reason: "relationship-changed", part: relsPath };
  }
  if (
    !contentTypesKept(
      contentTypes(before.parts.get(CONTENT_TYPES_PATH)),
      contentTypes(after.parts.get(CONTENT_TYPES_PATH))
    )
  ) {
    return { ok: false, reason: "part-changed", part: CONTENT_TYPES_PATH };
  }
  if (aroundTheStory(before) !== aroundTheStory(after)) {
    return { ok: false, reason: "part-changed", part: before.mainPartPath };
  }
  return { ok: true };
}

function storyKept(
  before: PMNode,
  after: PMNode,
  authorId: string,
  editableComments: EditableComments
): CommentOnlyVerdict {
  if (!changesOnlyComments(before, after)) return refused("body-changed");
  if (
    !commentIdentitiesKept(before, after) ||
    !commentAdditionsBy(before, after, authorId)
  ) {
    return refused("comment-author-forged");
  }
  if (
    !commentEditsOwned(before, after, {
      protection: "comments",
      authorId,
      editableComments,
    })
  ) {
    return refused("comment-not-owned");
  }
  return { ok: true };
}

/**
 * Whether the submitted file differs from the original in nothing but comments, every one of them
 * added, edited, moved, deleted, replied to or settled by the author with this identity.
 *
 * Every part of the package has to arrive as it left, save for the three a comment is written
 * across and the relationship and content type they are declared with; the document story itself
 * has to read as it did, comments aside. A comment carrying no recorded identity is everyone's to
 * edit here as it is in the editor (`schema/protection`), while a comment that appeared has to
 * carry this identity: a file can claim any author, and the editor's own hand in writing it is
 * not there to vouch for it. An identity already recorded is nobody's to rewrite.
 *
 * `editableComments: "all"` judges the file of an editor opened for a moderator, where every
 * comment was theirs to edit; an identity is nobody's to rewrite under either setting.
 *
 * Bytes that are not a readable docx are turned down the way opening one is, with a
 * `DocxImportError`, rather than being answered as a file that changed.
 */
export function onlyCommentsChangedBy(
  original: DocxBytes,
  submitted: DocxBytes,
  authorId: string,
  { editableComments = "own" }: { editableComments?: EditableComments } = {}
): CommentOnlyVerdict {
  const before = importDocx(original);
  const after = importDocx(submitted);
  const packaged = packageKept(before.session, after.session);
  return packaged.ok
    ? storyKept(before.doc, after.doc, authorId, editableComments)
    : packaged;
}
