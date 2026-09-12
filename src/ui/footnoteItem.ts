/**
 * The one control that puts a footnote in, which every menu over the text draws its row from.
 *
 * A right click is answered by one of two menus - the text menu, and the table menu wherever the
 * click lands in a cell with nothing selected (`editor/plugins/textContextMenu`) - and a footnote
 * goes into a table cell as readily as into a paragraph, a contract's note on a clause among them.
 * So what the row says and what it runs is written down here rather than once per menu, and the
 * command asked without a dispatch, which is the `canInsertFootnote` query itself, is what each
 * menu draws a dead row from.
 *
 * Whether the row is there at all is the surface's own declaration (`editor/stories/storyView`):
 * a note is one of the things a surface says it takes, since the part the surface is written back
 * as has to carry the note's own entry, and a menu asks that rather than asking which surface it
 * is drawn over.
 */

import type { LucideIcon } from "lucide-react";
import { Superscript } from "lucide-react";
import type { Command } from "prosemirror-state";
import { insertFootnote } from "../editor/commands/footnoteCommands";
import type { SurfaceCapabilities } from "../editor/stories/storyView";
import { modifierLabels } from "./shortcutLabels";

export interface FootnoteItem {
  readonly label: string;
  readonly icon: LucideIcon;
  /** The shortcut for the same action, for a menu that labels its rows with theirs */
  readonly hint: string;
  readonly command: Command;
}

/** The row a menu draws, and null for a surface that takes no note */
export function footnoteItem(takes: SurfaceCapabilities): FootnoteItem | null {
  if (!takes.has("note")) return null;
  const { mod, alt } = modifierLabels();
  return {
    label: "Insert footnote",
    icon: Superscript,
    hint: `${mod}${alt}F`,
    command: insertFootnote,
  };
}
