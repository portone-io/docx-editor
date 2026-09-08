/**
 * One convention for a value the whole document decides: the comments beside it, the notes under
 * it, the list markers drawn over it.
 *
 * Such a value is worked out by walking the document, and what it is worked out from changes only
 * when the document does. Held here, it is derived once per edit rather than once per read, and a
 * caret moving, a menu opening or a component rendering again all read back the value the last
 * edit left. What a caller writes is unchanged: `read` answers with a value either way.
 *
 * A projection derives from the document alone. Anything that also depends on the selection, the
 * protection or another plugin belongs in that plugin, which is asked again after every
 * transaction rather than after an edit alone.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  type EditorState,
  Plugin,
  PluginKey,
  type PluginSpec,
} from "prosemirror-state";

export interface DocumentProjection<T> {
  plugin: Plugin<T>;
  /**
   * The derived value. A state built without the plugin - one a consumer assembled itself, or one
   * a test made out of a bare document - is answered by deriving on the spot, so a public reader
   * standing on a projection keeps answering whoever asks it.
   */
  read(state: EditorState): T;
}

export function documentProjection<T>(
  name: string,
  derive: (doc: PMNode) => T,
  /**
   * The view props drawn from the derived value, which is how a projection that paints what it
   * derived - comment ranges over the text they were written for - stays one plugin. `this` in
   * such a prop is the plugin, so it reads the value back with `this.getState(state)`.
   */
  props?: PluginSpec<T>["props"]
): DocumentProjection<T> {
  const key = new PluginKey<T>(name);
  return {
    plugin: new Plugin<T>({
      key,
      state: {
        init: (_config, state) => derive(state.doc),
        apply: (tr, held) => (tr.docChanged ? derive(tr.doc) : held),
      },
      props,
    }),
    read: (state) => key.getState(state) ?? derive(state.doc),
  };
}
