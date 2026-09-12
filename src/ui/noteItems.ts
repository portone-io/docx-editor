/**
 * The controls that put a note in, which every menu over the text draws its rows from.
 *
 * A right click is answered by one of two menus - the text menu, and the table menu wherever the
 * click lands in a cell with nothing selected (`editor/plugins/textContextMenu`) - and a note goes
 * into a table cell as readily as into a paragraph, so each row is written down once here. The
 * command asked without a dispatch is what each menu draws a dead row from, and whether the rows
 * are there at all is the surface's own declaration (`editor/stories/storyView`).
 */

import type { LucideIcon } from "lucide-react";
import { ListEnd, Superscript } from "lucide-react";
import type { Command } from "prosemirror-state";
import { insertEndnote, insertFootnote } from "../editor/commands/noteCommands";
import { NOTE_KEYS } from "../editor/plugins/keymap";
import type { SurfaceCapabilities } from "../editor/stories/storyView";
import { modifierLabels } from "./shortcutLabels";

export interface NoteItem {
  readonly label: string;
  readonly icon: LucideIcon;
  /** The shortcut for the same action, for a menu that labels its rows with theirs */
  readonly hint: string;
  readonly command: Command;
}

/** The letter a binding ends on, which is what a row names after the modifiers */
function keyLetter(binding: string): string {
  return (binding.split("-").at(-1) ?? "").toUpperCase();
}

const NO_ITEMS: readonly NoteItem[] = [];

/** The rows a menu draws, and none for a surface that takes no note */
export function noteItems(takes: SurfaceCapabilities): readonly NoteItem[] {
  if (!takes.has("note")) return NO_ITEMS;
  const { mod, alt } = modifierLabels();
  // The key each kind is put in on is the keymap's own, so a row cannot promise one the editor
  // does not hold: an endnote is not on the same key everywhere (`editor/plugins/keymap`)
  return [
    {
      label: "Insert footnote",
      icon: Superscript,
      hint: `${mod}${alt}${keyLetter(NOTE_KEYS.footnote)}`,
      command: insertFootnote,
    },
    {
      label: "Insert endnote",
      icon: ListEnd,
      hint: `${mod}${alt}${keyLetter(NOTE_KEYS.endnote)}`,
      command: insertEndnote,
    },
  ];
}
