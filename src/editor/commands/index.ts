/**
 * The public surface for character formatting and list editing.
 *
 * Everything the built-in toolbar and the right click menu run is exported, so a consumer can
 * build controls of its own design or drive the same editing from a plugin. Following the
 * ProseMirror convention, calling a command without `dispatch` only reports whether it is
 * applicable right now, and `is*Active`, `can*` and `active*` are used to render button state.
 */

import type { Command } from "prosemirror-state";

export type EditorCommand = Command;

/** The paragraph styles the document defines, which the style picker offers */
export type { ParagraphStyleOption } from "../../docx/formatting";
export type { RunToggle } from "../../docx/runProps";
export type { LineSpacing, ParagraphAlign } from "../../model/format";
export type { ListKind } from "../../numbering/listTemplate";
/** The size an image is shown at, which the insert command asks for */
export type { ImageExtent } from "../../ooxml/image";
/**
 * What the editor as a whole may receive - everything, comments alone, or nothing - which a control
 * of your own reads before it offers an edit, the way the built-in ones do.
 */
export type {
  EditableComments,
  EditingProtection,
} from "../../schema/protection";
export { editingProtection } from "../../schema/protectionState";
/** The fonts to draw with where a document names one the machine does not have, which the font queries read */
export type { FontFallbacks } from "../../styles/fontStack";
/**
 * The document's own default formatting, which `documentFontNames` needs, and the width of one
 * line of body text on the paper the document names, which is what an image wider than the page
 * is shrunk to
 */
export {
  documentBodyWidthPx,
  documentDefaults,
  documentParagraphStyles,
} from "../documentStyles";
/**
 * The way an image file goes from outside into the document, which is what an image button of
 * your own stands on. `insertImageFiles` is all of it in one call, and is what the built-in
 * button runs; `readImageFile` and `fittedExtent` are for a caller that puts the image in itself,
 * the second being the rule by which an image wider than the body is shrunk to the body width.
 * Both take that width, which `documentBodyWidthPx` reads off the open document; left out, it is
 * A4's.
 */
export {
  fittedExtent,
  IMAGE_FILE_ACCEPT,
  imageFilesIn,
  insertImageFiles,
  readImageFile,
} from "../imageFiles";
export type { ImageToInsert } from "../insertImage";
export { canInsertImage, insertImage } from "../insertImage";
export type { TableSize } from "../insertTable";
export { canInsertTable, insertTable } from "../insertTable";
/**
 * The two breaks, which are what the keymap binds to Shift+Enter and Mod+Enter
 * (`editor/plugins/keymap`), so a control of your own can put in either.
 */
export { insertLineBreak, insertPageBreak } from "./breakCommands";
/**
 * Whether a command this package does not own would go through right now, the lock guard included.
 * The commands here need no such wrapper: each of them asks the guard before it answers. What does
 * is a command of a consumer's own, or one out of `prosemirror-commands`, since the predicates the
 * question takes are not exported anywhere.
 */
export { canRunCommand } from "./canRunCommand";
export type {
  CommentAuthor,
  DocumentComment,
  DocumentCommentReply,
  NewComment,
} from "./commentCommands";
export {
  addComment,
  addCommentReply,
  canAddComment,
  canEditComment,
  documentComments,
  removeComment,
  removeCommentReply,
  selectComment,
  setCommentResolved,
  updateComment,
  updateCommentReply,
} from "./commentCommands";
export type {
  ActiveFontFamily,
  ActiveFontSize,
} from "./formattingCommands";
export {
  activeFontFamily,
  activeFontSize,
  activeTextBackground,
  activeTextColor,
  canFormatText,
  documentFontNames,
  isBoldActive,
  isItalicActive,
  isStrikeActive,
  isUnderlineActive,
  setFontFamily,
  setFontSize,
  setTextBackground,
  setTextColor,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
} from "./formattingCommands";
/**
 * Undo and redo are the editor's own rather than prosemirror-history's, because a replayed lock
 * only passes the lock guard with the pass these two carry (`./historyCommands`). They also find
 * the history this editor keeps, which an import of the consumer's own copy of prosemirror-history
 * would not.
 */
export { redo, undo } from "./historyCommands";
export {
  canDecreaseIndent,
  canIncreaseIndent,
  decreaseIndent,
  increaseIndent,
} from "./indentCommands";
export type { ActiveLinkSpan } from "./linkCommands";
/**
 * The link commands, which the Cmd+K panel and the toolbar button are drawn from
 * (`ui/LinkPanel`). `activeLink` is the address to show, `canSetLink` whether a link may go on at
 * all - it is asked before an address has been typed - and `removeLink` reporting false is how a
 * control knows there is no link to take away.
 * `activeLinkSpan` is the one link the selection sits inside, stretch and address, which is what a
 * card standing by a link is placed and drawn from (`ui/LinkCard`).
 */
export {
  activeLink,
  activeLinkSpan,
  canSetLink,
  removeLink,
  setLink,
} from "./linkCommands";
export {
  activeListKind,
  decreaseListLevel,
  increaseListLevel,
  isInList,
  toggleBulletList,
  toggleNumberedList,
} from "./listCommands";
export type { SelectionLock } from "./lockCommands";
export {
  documentHasLocked,
  lockSelection,
  selectionLock,
  selectionTouchesLocked,
  unlockSelection,
} from "./lockCommands";
export type { DocumentNote } from "./noteQueries";
export { documentNotes } from "./noteQueries";
export type {
  ActiveParagraphAlign,
  ActiveParagraphStyle,
} from "./paragraphCommands";
export {
  activeParagraphAlign,
  activeParagraphStyle,
  canSetParagraphAlign,
  setParagraphAlign,
  setParagraphStyle,
} from "./paragraphCommands";
export {
  activeLineSpacing,
  canSetLineSpacing,
  SINGLE_LINE_SPACING,
  setLineSpacing,
} from "./spacingCommands";
/** Inserts the same editable DOCX tab as the built-in Tab binding. */
export { insertTab } from "./tabCommands";
