/**
 * Whether the panel that reads and writes a link's address stands open.
 *
 * Like the two context menus, this file only records that the panel was asked for; drawing it is
 * the React side's job (`ui/LinkPanel`). One state serves both ways in - Cmd+K
 * (`editor/plugins/keymap`) and the toolbar button - so the panel is the same panel however it was
 * opened, and it opens where the link is going rather than under whichever control asked for it.
 *
 * It closes on the document changing, which is what applying a link or taking one off does: the
 * panel is about the selection it opened over, and after an edit that selection has been answered.
 */

import type { Command, EditorState, Plugin } from "prosemirror-state";
import { canSetLink } from "../commands/linkCommands";
import { panelPlugin } from "./panelState";

/** The panel is about the selection it opened over, so standing open is all there is to hold */
const OPEN = true;

const panel = panelPlugin<typeof OPEN>({
  name: "docxEditorLinkPanel",
  isAnchor: (value): value is typeof OPEN => value === OPEN,
  // An edit answers the panel: the link went on, came off, or the text it was about moved
  onDocChange: "close",
  canOpen: (state) => canSetLink(state),
});

/** Whether the link panel is open */
export function isLinkPanelOpen(state: EditorState): boolean {
  return panel.anchor(state) === OPEN;
}

/**
 * Opens the panel. It reports false where a link could not go on anyway - a caret standing in no
 * link, a selection a lock leaves nothing of - so that Cmd+K there is a key this editor did not
 * take, and a control drawn from it is drawn dead. A panel already open is one more such place:
 * the button that opened it has nothing left to do while it stands.
 */
export const openLinkPanel: Command = (state, dispatch) =>
  !isLinkPanelOpen(state) && panel.open(OPEN)(state, dispatch);

/** Closes it. Called on applying, on Escape, and on a click outside */
export const closeLinkPanel: Command = panel.close;

export function linkPanel(): Plugin<typeof OPEN | null> {
  return panel.plugin;
}
