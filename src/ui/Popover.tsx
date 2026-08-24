/**
 * Toolbar-button popover. Keyboard activation moves focus inside and restores it on close;
 * pointer activation leaves the current focus unchanged.
 */

import type { LucideIcon } from "lucide-react";
import { type ReactNode, useCallback, useRef, useState } from "react";
import { editorClassNames } from "../styles/classNames";
import { usePanelUnderControl } from "./panelPlacement";
import { type RunSource, ToolbarButton } from "./ToolbarButton";
import { useDismiss } from "./useDismiss";

/** Menus provide their own role; dialogs use the popover wrapper. */
export type PopoverPanel = "menu" | "dialog";

export interface PanelHandle {
  close: () => void;
  /** Whether keyboard activation requested focus inside the panel. */
  takeFocus: boolean;
}

export interface PopoverProps {
  label: string;
  icon: LucideIcon;
  indicatorColor?: string;
  panel: PopoverPanel;
  disabled?: boolean;
  children: (handle: PanelHandle) => ReactNode;
}

export function Popover({
  label,
  icon,
  indicatorColor,
  panel: kind,
  disabled,
  children,
}: PopoverProps) {
  const [open, setOpen] = useState(false);
  const box = useRef<HTMLDivElement | null>(null);
  const button = useRef<HTMLButtonElement | null>(null);
  const panel = useRef<HTMLDivElement | null>(null);
  // Restore focus only when this popover moved it in the first place.
  const holdsFocus = useRef(false);
  const close = useCallback(() => {
    setOpen(false);
    if (holdsFocus.current) button.current?.focus({ preventScroll: true });
    holdsFocus.current = false;
  }, []);
  useDismiss(box, open, close);
  const position = usePanelUnderControl(panel, button, open);

  const toggle = (source: RunSource) => {
    if (open) {
      close();
      return;
    }
    holdsFocus.current = source === "keyboard";
    setOpen(true);
  };

  return (
    <div className={editorClassNames.popoverBox} ref={box}>
      <ToolbarButton
        ref={button}
        label={label}
        icon={icon}
        indicatorColor={indicatorColor}
        popup={kind}
        expanded={open}
        disabled={disabled}
        onRun={toggle}
      />
      {open && (
        <div
          ref={panel}
          className={editorClassNames.popover}
          role={kind === "dialog" ? "dialog" : undefined}
          aria-label={kind === "dialog" ? label : undefined}
          style={{
            left: position?.left,
            top: position?.top,
            // Hide until the panel has been measured.
            visibility: position ? undefined : "hidden",
          }}
        >
          {children({ close, takeFocus: holdsFocus.current })}
        </div>
      )}
    </div>
  );
}
