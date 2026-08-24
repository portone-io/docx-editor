/** Choice row with a fixed check slot for stable label alignment. */

import { Check, type LucideIcon } from "lucide-react";
import { editorClassNames } from "../styles/classNames";
import { ICON_SIZE } from "./ToolbarButton";

export interface MenuChoiceProps {
  label: string;
  icon?: LucideIcon;
  role: "menuitemradio" | "menuitemcheckbox";
  checked: boolean;
  onChoose: () => void;
}

export function MenuChoice({
  label,
  icon: Icon,
  role,
  checked,
  onChoose,
}: MenuChoiceProps) {
  return (
    <button
      type="button"
      role={role}
      className={editorClassNames.menuItem}
      aria-checked={checked}
      onMouseDown={(event) => event.preventDefault()}
      onClick={onChoose}
    >
      <span className={editorClassNames.menuCheck}>
        {checked && <Check size={ICON_SIZE} aria-hidden="true" />}
      </span>
      {Icon && <Icon size={ICON_SIZE} aria-hidden="true" />}
      {label}
    </button>
  );
}
