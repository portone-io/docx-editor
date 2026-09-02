import { R_NS } from "../../ooxml/xml";

export const COMMENTS_REL_TYPE = `${R_NS}/comments`;
export const COMMENTS_EXTENDED_REL_TYPE =
  "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
export const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
export const COMMENTS_EXTENDED_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
export const CONTENT_TYPES_PATH = "[Content_Types].xml";
export const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
export const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";
export const MC_NS =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const PEOPLE_REL_TYPE =
  "http://schemas.microsoft.com/office/2011/relationships/people";
export const PEOPLE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml";
/**
 * The `w15:providerId` under which this editor records a comment author's identity. An identity
 * another provider recorded, Word's directory above all, is one this editor cannot vouch for and
 * is read as none.
 */
export const COMMENT_AUTHOR_PROVIDER = "portone-docx-editor";
