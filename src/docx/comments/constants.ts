import { NAMESPACES } from "../../ooxml/names";
import { R_NS } from "../../ooxml/xml";

export const COMMENTS_REL_TYPE = `${R_NS}/comments`;
export const COMMENTS_EXTENDED_REL_TYPE =
  "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
export const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
export const COMMENTS_EXTENDED_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
export const W14_NS = NAMESPACES.w14;
export const W15_NS = NAMESPACES.w15;
export const MC_NS = NAMESPACES.mc;
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
