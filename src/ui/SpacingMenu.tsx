import type { Command, EditorState } from "prosemirror-state";
import { useRef } from "react";
import {
  activeLineSpacing,
  setLineSpacing,
} from "../editor/commands/spacingCommands";
import type { LineSpacing } from "../model/format";
import { editorClassNames } from "../styles/classNames";
import {
  DEFAULT_LINE_SPACINGS,
  type LineSpacingOption,
} from "../styles/lineSpacings";
import { MenuChoice } from "./MenuChoice";
import type { RunCommand } from "./runCommand";
import { useMenuKeyboard } from "./useMenuKeyboard";

function isCurrentChoice(
  current: LineSpacing | null,
  { spacing }: LineSpacingOption
): boolean {
  return (
    current?.rule === "auto" &&
    spacing.rule === "auto" &&
    current.lines === spacing.lines
  );
}

export interface SpacingMenuProps {
  state: EditorState;
  lineSpacings?: readonly LineSpacingOption[];
  run: RunCommand;
  close: () => void;
  takeFocus: boolean;
}

export function SpacingMenu({
  state,
  lineSpacings = DEFAULT_LINE_SPACINGS,
  run,
  close,
  takeFocus,
}: SpacingMenuProps) {
  const current = activeLineSpacing(state);
  const menu = useRef<HTMLDivElement | null>(null);
  const keys = useMenuKeyboard({ menu, onClose: close, takeFocus });

  const choose = (command: Command) => {
    run(command);
    close();
  };

  return (
    <div
      ref={menu}
      className={editorClassNames.menuList}
      role="menu"
      aria-label="Line and paragraph spacing"
      {...keys}
    >
      {lineSpacings.map((option) => (
        <MenuChoice
          key={option.label}
          label={option.label}
          role="menuitemradio"
          checked={isCurrentChoice(current, option)}
          onChoose={() => choose(setLineSpacing(option.spacing))}
        />
      ))}
    </div>
  );
}
