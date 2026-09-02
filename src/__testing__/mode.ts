import type { DocxEditorMode } from "../DocxEditor";
import type { CommentAuthor } from "../editor/commands/comments/model";

/** The identity the tests write comments under */
export const AUTHOR: CommentAuthor = { id: "tester", name: "Tester" };

/** The mode an editing test mounts in, which is what every test mounted in before modes named an author */
export const EDITING: DocxEditorMode = { kind: "edit", author: AUTHOR };
