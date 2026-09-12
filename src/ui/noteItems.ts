/**
 * The controls that put a note in, which every menu over the text draws its rows from.
 *
 * A right click is answered by one of two menus - the text menu, and the table menu wherever the
 * click lands in a cell with nothing selected (`editor/plugins/textContextMenu`) - and a note goes
 * into a table cell as readily as into a paragraph, a contract's note on a clause among them. So
 * what each row says and what it runs is written down here rather than once per menu, and the
 * command asked without a dispatch, which is the `canInsert` query itself, is what each menu draws
 * a dead row from.
 *
 * There is a row per kind because there is a command per kind (`editor/commands/noteCommands`),
 * and the two differ in where the note is drawn rather than in anything a menu has to know: a
 * footnote stands at the foot of the page that calls it and an endnote after the last paragraph.
 *
 * Whether the rows are there at all is the surface's own declaration (`editor/stories/storyView`):
 * a note is one of the things a surface says it takes, since the part the surface is written back
 * as has to carry the note's own entry, and a menu asks that rather than asking which surface it
 * is drawn over.
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
