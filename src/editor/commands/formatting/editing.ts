/** The character formatting commands the package exports, each built from the property it edits. */

import type { Command, EditorState } from "prosemirror-state";
import { editShut, openStretches } from "../../../schema/guards";
import { runPropertyCommands } from "./propertyCommands";
import { textPieces } from "./shared";

/**
 * Whether character formatting reaches anything where the selection stands, which is what a
 * control offering it is drawn from.
 *
 * A selection running over a lock keeps its say, because the locked pieces are left out rather
 * than the whole edit refused; only a selection the lock leaves nothing of has nothing to format.
 * A caret formats no text of its own - it stages the formatting for the text typed next - so what
 * settles it is whether that text could go in at all.
 */
export function canFormatText(state: EditorState): boolean {
  if (state.selection.empty) {
    return !editShut(state, { kind: "insert", at: state.selection.from });
  }
  return openStretches(state, textPieces(state), "mark").length > 0;
}

const bold = runPropertyCommands("bold");
const italic = runPropertyCommands("italic");
const underline = runPropertyCommands("underline");
const strike = runPropertyCommands("strike");
const fontSize = runPropertyCommands("fontSizePt");
const fontFamily = runPropertyCommands("fontFamily");
const textColor = runPropertyCommands("color");
const textBackground = runPropertyCommands("background");

export const toggleBold: Command = bold.toggle(true);
export const toggleItalic: Command = italic.toggle(true);
/** Switching the underline on writes a single one; a run already underlined keeps the kind it has */
export const toggleUnderline: Command = underline.toggle("single");
export const toggleStrike: Command = strike.toggle(true);

/** Sets the font size in points. Null withdraws the setting and falls back to the document default */
export function setFontSize(pt: number | null): Command {
  return fontSize.set(pt);
}

/** Sets the font by name. Null withdraws the setting and falls back to the document default font */
export function setFontFamily(name: string | null): Command {
  return fontFamily.set(name);
}

/** Sets the text color as `#RRGGBB`. Null withdraws the color setting */
export function setTextColor(hex: string | null): Command {
  return textColor.set(hex);
}

/**
 * Sets the text background color as `#RRGGBB`. Null withdraws the background.
 * A highlight (`w:highlight`) written by an older document is removed along with it at that spot.
 */
export function setTextBackground(hex: string | null): Command {
  return textBackground.set(hex);
}

export function isBoldActive(state: EditorState): boolean {
  return bold.isActive(state);
}

export function isItalicActive(state: EditorState): boolean {
  return italic.isActive(state);
}

export function isUnderlineActive(state: EditorState): boolean {
  return underline.isActive(state);
}

export function isStrikeActive(state: EditorState): boolean {
  return strike.isActive(state);
}
