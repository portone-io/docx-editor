/**
 * Tiny packages carrying Word comments, shared by the tests that read comment anchors and the ones
 * that edit the text under them.
 */

import { unzipSync, zipSync } from "fflate";
import { makeDocx } from "./docx";

const encoder = new TextEncoder();
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CONTENT_TYPES_NS =
  "http://schemas.openxmlformats.org/package/2006/content-types";

/** One run of text, written the way the writer writes one */
export function run(text: string): string {
  return `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
}

/** The text a comment marks, between its range markers and followed by its reference */
export function commentedRun(id: string, text: string): string {
  return (
    `<w:commentRangeStart w:id="${id}"/>${run(text)}` +
    `<w:commentRangeEnd w:id="${id}"/>` +
    `<w:r><w:commentReference w:id="${id}"/></w:r>`
  );
}

export interface CommentEntry {
  id: string;
  /** What the comment says, as one paragraph of plain text */
  text: string;
  author?: string;
  initials?: string;
  date?: string;
}

/** The attributes an entry names, left out where it names none, the way Word leaves them out */
function commentAttrs({ id, author, initials, date }: CommentEntry): string {
  return [
    ["id", id],
    ["author", author],
    ["initials", initials],
    ["date", date],
  ]
    .filter(([, value]) => value !== undefined)
    .map(([name, value]) => ` w:${name}="${value}"`)
    .join("");
}

function commentsPart(comments: readonly CommentEntry[]): string {
  const entries = comments.map(
    (comment) =>
      `<w:comment${commentAttrs(comment)}>` +
      `<w:p>${run(comment.text)}</w:p></w:comment>`
  );
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:comments xmlns:w="${W_NS}">${entries.join("")}</w:comments>`
  );
}

/** A package whose body is this markup, with a comments part holding these entries */
export function commentedDocx(
  body: string,
  comments: readonly CommentEntry[]
): Uint8Array {
  const parts = unzipSync(makeDocx(body));
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId5" Target="comments.xml" Type="${REL_BASE}/comments"/>` +
      "</Relationships>"
  );
  parts["word/comments.xml"] = encoder.encode(commentsPart(comments));
  parts["[Content_Types].xml"] = encoder.encode(
    `<Types xmlns="${CONTENT_TYPES_NS}">` +
      '<Override PartName="/word/document.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}
