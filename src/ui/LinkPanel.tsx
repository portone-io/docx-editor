/** Edits or removes the link on the current selection. */

import { Check, Link2 } from "lucide-react";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import {
  type KeyboardEvent,
  useCallback,
  useEffect,
  useRef,
  useState,
} from "react";
import {
  activeLink,
  removeLink,
  setLink,
} from "../editor/commands/linkCommands";
import { closeLinkPanel } from "../editor/plugins/linkPanel";
import { editorClassNames } from "../styles/classNames";
import { pointBelowPos, usePanelAtPoint } from "./panelPlacement";
import { commandRunner } from "./runCommand";
import { ICON_SIZE } from "./ToolbarButton";
import { tooltipAttribute } from "./Tooltip";
import { useDismiss } from "./useDismiss";

const FOCUSABLE = "input,button:not(:disabled)";

export interface LinkPanelProps {
  view: EditorView;
  state: EditorState;
}

export function LinkPanel({ view, state }: LinkPanelProps) {
  const box = useRef<HTMLDivElement | null>(null);
  const field = useRef<HTMLInputElement | null>(null);
  const current = activeLink(state);
  const [address, setAddress] = useState(current ?? "");
  const placement = usePanelAtPoint(
    box,
    pointBelowPos(view, state.selection.head)
  );
  const run = commandRunner(view);
  const close = useCallback(() => {
    closeLinkPanel(view.state, (tr) => view.dispatch(tr));
    // The dialog owns focus while open, so dismissal returns it to the editor.
    view.focus();
  }, [view]);
  useDismiss(box, true, close);

  // Wait for placement so an invisible dialog cannot take focus.
  useEffect(() => {
    const input = field.current;
    if (!input || !placement || input === document.activeElement) return;
    input.focus({ preventScroll: true });
    // The address already there is the one being changed, so typing replaces it
    input.select();
  }, [placement]);

  // An emptied address remains the form's compact way of taking a link off.
  const emptied = address.trim() === "";
  const applies = emptied ? removeLink(state) : setLink(address)(state);
  const apply = () => {
    run(emptied ? removeLink : setLink(address));
    close();
  };

  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key !== "Tab") return;
    const controls = Array.from(
      box.current?.querySelectorAll<HTMLElement>(FOCUSABLE) ?? []
    );
    if (controls.length === 0) return;
    const at = controls.findIndex(
      (control) => control === document.activeElement
    );
    const step = event.shiftKey ? -1 : 1;
    const next = controls[(at + step + controls.length) % controls.length];
    next.focus();
    event.preventDefault();
  };

  return (
    <div
      ref={box}
      className={editorClassNames.linkPanel}
      role="dialog"
      aria-label="Link"
      onKeyDown={onKeyDown}
      style={{
        left: placement?.left,
        top: placement?.top,
        // Hide until the dialog has been measured.
        visibility: placement ? undefined : "hidden",
      }}
    >
      <div className={editorClassNames.linkPanelRow}>
        <Link2
          size={ICON_SIZE}
          aria-hidden="true"
          className={editorClassNames.linkPanelIcon}
        />
        <input
          ref={field}
          type="text"
          className={editorClassNames.linkAddress}
          aria-label="Address"
          value={address}
          onChange={(event) => setAddress(event.target.value)}
          onKeyDown={(event) => {
            if (event.key !== "Enter") return;
            event.preventDefault();
            apply();
          }}
        />
        <button
          type="button"
          className={editorClassNames.toolbarButton}
          disabled={!applies}
          onClick={apply}
          aria-label="Apply"
          {...tooltipAttribute("Apply")}
        >
          <Check size={ICON_SIZE} aria-hidden="true" />
        </button>
      </div>
    </div>
  );
}
