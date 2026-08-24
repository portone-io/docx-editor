/** Shows the destination and available actions for the link under the selection. */

import { ExternalLink, Link2Off, Pencil } from "lucide-react";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useCallback, useEffect, useRef, useState } from "react";
import {
  type ActiveLinkSpan,
  removeLink,
} from "../editor/commands/linkCommands";
import { openLinkPanel } from "../editor/plugins/linkPanel";
import { editorClassNames } from "../styles/classNames";
import { openAddress, openable } from "./openAddress";
import {
  pointBelowPos,
  type ScreenPoint,
  usePanelAtPoint,
} from "./panelPlacement";
import { commandRunner } from "./runCommand";
import { ICON_SIZE } from "./ToolbarButton";
import { Tooltip, tooltipAttribute } from "./Tooltip";

/** Keeps the card anchored to the link's first line while the caret moves within it. */
function useLinkPoint(view: EditorView, at: number): ScreenPoint {
  const [point, setPoint] = useState(() => pointBelowPos(view, at));

  useEffect(() => {
    const measure = () =>
      setPoint((was) => {
        const now = pointBelowPos(view, at);
        return was.clientX === now.clientX && was.clientY === now.clientY
          ? was
          : now;
      });
    measure();
    document.addEventListener("scroll", measure, true);
    window.addEventListener("resize", measure);
    return () => {
      document.removeEventListener("scroll", measure, true);
      window.removeEventListener("resize", measure);
    };
  }, [view, at]);

  return point;
}

function useEscape(onEscape: () => void): void {
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onEscape();
    };
    document.addEventListener("keydown", onKeyDown, true);
    return () => document.removeEventListener("keydown", onKeyDown, true);
  }, [onEscape]);
}

export interface LinkCardProps {
  view: EditorView;
  state: EditorState;
  link: ActiveLinkSpan;
  readOnly?: boolean;
}

export function LinkCard({
  view,
  state,
  link,
  readOnly = false,
}: LinkCardProps) {
  const box = useRef<HTMLDivElement | null>(null);
  const [hidden, setHidden] = useState(false);
  const placement = usePanelAtPoint(box, useLinkPoint(view, link.from));
  const run = commandRunner(view);
  const hide = useCallback(() => {
    if (box.current?.contains(document.activeElement)) view.focus();
    setHidden(true);
  }, [view]);
  useEscape(hide);

  if (hidden) return null;

  const href = link.href;
  return (
    <div
      ref={box}
      className={editorClassNames.linkCard}
      role="group"
      aria-label="Link"
      style={{
        left: placement?.left,
        top: placement?.top,
        // Hide until the card has been measured.
        visibility: placement ? undefined : "hidden",
      }}
    >
      {href === null ? (
        <span className={editorClassNames.linkCardNote}>
          This link points at a place in the document rather than at an address.
        </span>
      ) : (
        // Keep the tooltip outside the address's clipped overflow.
        <Tooltip label={href}>
          <span className={editorClassNames.linkCardAddress}>{href}</span>
        </Tooltip>
      )}
      {href !== null && (
        <button
          type="button"
          className={editorClassNames.toolbarButton}
          disabled={!openable(href)}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => openAddress(href)}
          aria-label="Open"
          {...tooltipAttribute("Open")}
        >
          <ExternalLink size={ICON_SIZE} aria-hidden="true" />
        </button>
      )}
      {!readOnly && openLinkPanel(state) && (
        <button
          type="button"
          className={editorClassNames.toolbarButton}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => run(openLinkPanel)}
          aria-label="Edit"
          {...tooltipAttribute("Edit")}
        >
          <Pencil size={ICON_SIZE} aria-hidden="true" />
        </button>
      )}
      {!readOnly && removeLink(state) && (
        <button
          type="button"
          className={editorClassNames.toolbarButton}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => run(removeLink)}
          aria-label="Remove"
          {...tooltipAttribute("Remove")}
        >
          <Link2Off size={ICON_SIZE} aria-hidden="true" />
        </button>
      )}
    </div>
  );
}
