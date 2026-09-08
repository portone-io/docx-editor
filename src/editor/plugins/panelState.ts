/**
 * One convention for every panel the editor opens over a spot in the document: a link panel, a
 * right click menu, the comment composer.
 *
 * Each of them holds the same thing - where it stands, or nothing when it is shut - and answers the
 * same three questions: open it here, close it, where does it stand. What differs between them is
 * what "here" is and what an edit or a change of protection does to a panel already open, and those
 * are the options below rather than four copies of a plugin.
 *
 * The anchor is whatever the panel is about: a point on the screen for a menu, a stretch of the
 * document for the composer. It is read back out of the state as it was put in, so a panel that
 * needs nothing but "open" carries an anchor that says just that.
 *
 * A panel that reads other plugins in `canOpen` or `closeWhen` must be registered after them:
 * the state a plugin is handed while the transaction is being applied holds the fields of the
 * plugins ahead of it and nothing of the ones behind it.
 */

import {
  type Command,
  type EditorState,
  Plugin,
  PluginKey,
  type PluginSpec,
  type Transaction,
} from "prosemirror-state";

/**
 * What a document change does to an open panel: shut it, or move the anchor along with the change.
 * Answering null from the mapping shuts it, which is what a stretch the edit swallowed comes to.
 */
export type PanelDocChange<Anchor> =
  | "close"
  | ((anchor: Anchor, tr: Transaction, state: EditorState) => Anchor | null);

export interface PanelOptions<Anchor> {
  /** The plugin key's name, which is what a state debugger and a duplicate registration report */
  name: string;
  /**
   * Whether this is an anchor of this panel's. The opening goes through the transaction metadata,
   * which anything may write, so what is read back is judged rather than trusted
   */
  isAnchor(value: unknown): value is Anchor;
  onDocChange: PanelDocChange<Anchor>;
  /** Whether the panel may open here. `open` reports false where it says no */
  canOpen?(state: EditorState, anchor: Anchor): boolean;
  /**
   * Whether the panel must shut where the state now stands, asked after every transaction rather
   * than after an edit alone: a protection switch changes no text (`./documentProtection`), so
   * `onDocChange` never sees it, and a panel that only a commenter is offered has to go with it.
   */
  closeWhen?(state: EditorState, anchor: Anchor): boolean;
  /**
   * The view props the panel needs of its own, which is how a panel opened by a DOM event - a
   * right click - stays one plugin. What that handler does with the event is the panel's business;
   * all this factory asks is that the opening go through `opening` so the state stays this one.
   */
  props?: PluginSpec<Anchor | null>["props"];
}

export interface PanelPlugin<Anchor> {
  plugin: Plugin<Anchor | null>;
  /**
   * Opens the panel over this anchor, or moves an open one there. False where `canOpen` says the
   * panel has nothing to do here, so that a control drawn from it is drawn dead and a key bound to
   * it is a key this editor did not take.
   */
  open(anchor: Anchor): Command;
  /**
   * The same opening recorded on a transaction already being built, for an opener that moves the
   * selection in the same step. `canOpen` is not asked here: an opener holding the event that
   * asked - a right click landing where the menu belongs - has decided that already.
   */
  opening(tr: Transaction, anchor: Anchor): Transaction;
  /** Closes it. False where it stands closed already, so a key press falls through to what is behind it */
  close: Command;
  /** Where the panel stands. Null when it is shut, and for a state built without the plugin */
  anchor(state: EditorState): Anchor | null;
}

/** The one value that closes a panel. Fixed, so it can never be mistaken for anchor data */
const CLOSE = Symbol("panelClose");

export function panelPlugin<Anchor>({
  name,
  isAnchor,
  onDocChange,
  canOpen,
  closeWhen,
  props,
}: PanelOptions<Anchor>): PanelPlugin<Anchor> {
  const key = new PluginKey<Anchor | null>(name);
  const anchor = (state: EditorState): Anchor | null =>
    key.getState(state) ?? null;

  /** Where the panel stands after this transaction, before the state it lands in is asked about it */
  const moved = (
    tr: Transaction,
    held: Anchor | null,
    next: EditorState
  ): Anchor | null => {
    const meta: unknown = tr.getMeta(key);
    if (meta === CLOSE) return null;
    if (isAnchor(meta)) return meta;
    if (held === null || !tr.docChanged) return held;
    return onDocChange === "close" ? null : onDocChange(held, tr, next);
  };

  const opening = (tr: Transaction, over: Anchor): Transaction =>
    tr.setMeta(key, over);

  return {
    plugin: new Plugin<Anchor | null>({
      key,
      state: {
        init: () => null,
        apply: (tr, held, _before, next) => {
          const standing = moved(tr, held, next);
          if (standing === null) return null;
          return closeWhen?.(next, standing) ? null : standing;
        },
      },
      props,
    }),
    open:
      (over) =>
      (state, dispatch): boolean => {
        if (canOpen && !canOpen(state, over)) return false;
        dispatch?.(opening(state.tr, over));
        return true;
      },
    opening,
    close: (state, dispatch) => {
      if (anchor(state) === null) return false;
      dispatch?.(state.tr.setMeta(key, CLOSE));
      return true;
    },
    anchor,
  };
}
