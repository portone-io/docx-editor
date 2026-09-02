/**
 * Rejects transactions disallowed by the guard in `schema/locks`, which answers for the locks the
 * document carries and for the protection the editor runs under (`schema/protection`) alike. A
 * rejected IME edit also ends the browser composition on the next frame so ProseMirror cannot
 * remain stuck in composing state.
 */

import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { transactionAllowed } from "../../schema/locks";

/**
 * Ends the composition the refusal has already broken, and redraws where it stood.
 *
 * Handing ProseMirror stored marks is the door it ends a composition through of its own accord
 * (`updateStateInner`), and it redraws whatever the composition left dirty in the same pass. The
 * marks handed over are the ones already in force where the caret stands, which is the value a
 * character typed there would have taken anyway, so nothing else about the state moves.
 */
function endComposition(view: EditorView): void {
  if (!view.composing) return;
  const marks = view.state.storedMarks ?? view.state.selection.$from.marks();
  view.dispatch(
    view.state.tr.setStoredMarks(marks).setMeta("addToHistory", false)
  );
}

export function lockedContent(): Plugin {
  let live: EditorView | null = null;
  let frame = 0;

  /** The refusal is answered on the frame after it, so the browser is done with the event first */
  const afterRefusal = () => {
    if (live?.composing !== true || frame !== 0) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      if (live) endComposition(live);
    });
  };

  return new Plugin({
    view(view) {
      live = view;
      return {
        destroy() {
          if (frame !== 0) cancelAnimationFrame(frame);
          frame = 0;
          live = null;
        },
      };
    },
    filterTransaction(tr, state) {
      const allowed = transactionAllowed(tr, state);
      if (!allowed) afterRefusal();
      return allowed;
    },
  });
}
