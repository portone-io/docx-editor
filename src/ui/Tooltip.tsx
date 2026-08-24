/** CSS-rendered tooltip shown on hover or keyboard focus. */

import { type ReactNode, useEffect, useState } from "react";
import { editorAttributes, editorClassNames } from "../styles/classNames";

/** Adds tooltip text to an element that can render its own pseudo-element. */
export function tooltipAttribute(label: string) {
  return { [editorAttributes.tooltip]: label };
}

/** Hides tooltips after Escape until pointer or focus movement, as required by WCAG 1.4.13. */
export function useTooltipsHushed(): boolean {
  const [hushed, setHushed] = useState(false);

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setHushed(true);
    };
    document.addEventListener("keydown", onKeyDown, true);
    return () => document.removeEventListener("keydown", onKeyDown, true);
  }, []);

  useEffect(() => {
    if (!hushed) return;
    const wake = () => setHushed(false);
    document.addEventListener("pointermove", wake, true);
    document.addEventListener("focusin", wake, true);
    return () => {
      document.removeEventListener("pointermove", wake, true);
      document.removeEventListener("focusin", wake, true);
    };
  }, [hushed]);

  return hushed;
}

export function hushedAttribute(hushed: boolean) {
  return hushed ? { [editorAttributes.tooltipsHushed]: "" } : {};
}

export interface TooltipProps {
  /** Text shown by the tooltip. */
  label: string;
  children: ReactNode;
}

export function Tooltip({ label, children }: TooltipProps) {
  return (
    <span
      className={editorClassNames.tooltipWrapper}
      {...tooltipAttribute(label)}
    >
      {children}
    </span>
  );
}
