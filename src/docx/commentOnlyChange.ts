/**
 * Whether a file handed back changed in nothing but comments one author is allowed to have made.
 *
 * The editor's own refusal under a comment protection (`schema/protection`) is a courtesy to the
 * user: the browser holds the file, so the rule is held again here, where a server takes it in.
 * The judgement is over the package rather than over the document alone, since a submission is
 * free to rewrite anything the story does not carry - the paper the document is written on, its
 * styles, its headers - and a document comparison would see none of it.
 */

import type { XmlParser } from "../ooxml/xml";
import type { EditableComments } from "../schema/protection";
import { commentsPolicy } from "./comments/policy";
import type { DocxBytes } from "./importDocx";
import { verifyChange } from "./protectionPolicy";

/**
 * Why a file is not the one it claims to be. `part-changed`, `relationship-changed` and
 * `comment-markup-rejected` name the part they were reached over; the other three are about the
 * document story itself.
 */
export type CommentOnlyVerdict =
  | { ok: true }
  | {
      ok: false;
      reason: "body-changed" | "comment-not-owned" | "comment-author-forged";
    }
  | {
      ok: false;
      reason:
        | "part-changed"
        | "relationship-changed"
        | "comment-markup-rejected";
      part: string;
    };

/**
 * Whether the submitted file differs from the original in nothing but comments, every one of them
 * added, edited, moved, deleted, replied to or settled by the author with this identity.
 *
 * Every part of the package has to arrive as it left, save for the three a comment is written
 * across and the relationship and content type they are declared with; the document story itself
 * has to read as it did, comments aside. Those three parts are read entry by entry instead
 * (`./comments/policy`), since a comment edit is free to rewrite them and something has to say
 * what it may have written there. A comment carrying no recorded identity is everyone's to
 * edit here as it is in the editor (`schema/protection`), while a comment that appeared has to
 * carry this identity: a file can claim any author, and the editor's own hand in writing it is
 * not there to vouch for it. An identity already recorded is nobody's to rewrite.
 *
 * `editableComments: "all"` judges the file of an editor opened for a moderator, where every
 * comment was theirs to edit; an identity is nobody's to rewrite under either setting.
 * `xmlParser` names the parser both files are read through, for a runtime that holds no
 * `DOMParser` global of its own.
 *
 * Bytes that are not a readable docx are turned down the way opening one is, with a
 * `DocxImportError`, rather than being answered as a file that changed.
 */
export function onlyCommentsChangedBy(
  original: DocxBytes,
  submitted: DocxBytes,
  authorId: string,
  {
    editableComments = "own",
    xmlParser,
  }: { editableComments?: EditableComments; xmlParser?: XmlParser } = {}
): CommentOnlyVerdict {
  return verifyChange(commentsPolicy, original, submitted, authorId, {
    editableComments,
    xmlParser,
  });
}
