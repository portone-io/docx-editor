/** A toolbar button whose pointer activation preserves the editor selection and IME state. */

import type { LucideIcon } from "lucide-react";
import { forwardRef } from "react";
import { editorClassNames } from "../styles/classNames";
import { tooltipAttribute } from "./Tooltip";

export const ICON_SIZE = 16;

/** Keyboard-generated clicks have `detail === 0`; panels use the source to decide focus. */
export type RunSource = "pointer" | "keyboard";

export interface ToolbarButtonProps {
  /** Shared accessible name and tooltip text. */
  label: string;
  icon: LucideIcon;
  pressed?: boolean;
  /** Open state for disclosure buttons. */
  expanded?: boolean;
  popup?: "menu" | "dialog";
  /** Optional color bar drawn under the icon. */
  indicatorColor?: string;
  disabled?: boolean;
  onRun: (source: RunSource) => void;
}

/**
 * `forwardRef` rather than a `ref` prop: React 18 hands a function component no `ref`
 * in its props, so the button would never reach the popovers that focus it.
 */
export const ToolbarButton = forwardRef<HTMLButtonElement, ToolbarButtonProps>(
  function ToolbarButton(
    {
      label,
      icon: Icon,
      pressed,
      expanded,
      popup,
      indicatorColor,
      disabled,
      onRun,
    },
    ref
  ) {
    return (
      <button
        ref={ref}
        type="button"
        className={editorClassNames.toolbarButton}
        aria-label={label}
        {...tooltipAttribute(label)}
        aria-pressed={pressed}
        aria-expanded={expanded}
        aria-haspopup={popup}
        disabled={disabled}
        onMouseDown={(event) => event.preventDefault()}
        onClick={(event) => onRun(event.detail > 0 ? "pointer" : "keyboard")}
      >
        {indicatorColor ? (
          <span
            className={editorClassNames.toolbarColorIcon}
            aria-hidden="true"
          >
            <Icon size={ICON_SIZE} />
            <span
              className={editorClassNames.toolbarColorIndicator}
              style={{ backgroundColor: indicatorColor }}
            />
          </span>
        ) : (
          <Icon size={ICON_SIZE} aria-hidden="true" />
        )}
      </button>
    );
  }
);
