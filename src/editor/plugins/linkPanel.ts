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

import {
  type Command,
  type EditorState,
  Plugin,
  PluginKey,
} from "prosemirror-state";
import { canSetLink } from "../commands/linkCommands";

const panelKey = new PluginKey<boolean>("docxEditorLinkPanel");

/** Whether the link panel is open */
export function isLinkPanelOpen(state: EditorState): boolean {
  return panelKey.getState(state) === true;
}

/**
 * Opens the panel. It reports false where a link could not go on anyway - a caret standing in no
 * link, a selection a lock leaves nothing of - so that Cmd+K there is a key this editor did not
 * take, and a control drawn from it is drawn dead.
 */
export const openLinkPanel: Command = (state, dispatch) => {
  if (isLinkPanelOpen(state) || !canSetLink(state)) return false;
  if (dispatch) dispatch(state.tr.setMeta(panelKey, true));
  return true;
};

/** Closes it. Called on applying, on Escape, and on a click outside */
export const closeLinkPanel: Command = (state, dispatch) => {
  if (!isLinkPanelOpen(state)) return false;
  if (dispatch) dispatch(state.tr.setMeta(panelKey, false));
  return true;
};

export function linkPanel(): Plugin<boolean> {
  return new Plugin<boolean>({
    key: panelKey,
    state: {
      init: () => false,
      apply: (tr, open) => {
        const meta: unknown = tr.getMeta(panelKey);
        if (typeof meta === "boolean") return meta;
        // An edit answers the panel: the link went on, came off, or the text it was about moved
        return tr.docChanged ? false : open;
      },
    },
  });
}
