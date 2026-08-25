/** Menu for choosing the current paragraph alignment. */

import {
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  type LucideIcon,
} from "lucide-react";
import type { EditorState } from "prosemirror-state";
import { type ReactElement, useRef } from "react";
import {
  type ActiveParagraphAlign,
  activeParagraphAlign,
  setParagraphAlign,
} from "../editor/commands/paragraphCommands";
import type { ParagraphAlign } from "../model/format";
import { editorClassNames } from "../styles/classNames";
import { MenuChoice } from "./MenuChoice";
import type { RunCommand } from "./runCommand";
import { useMenuKeyboard } from "./useMenuKeyboard";

interface AlignChoice {
  label: string;
  icon: LucideIcon;
  align: ParagraphAlign;
}

const ALIGNMENTS: readonly AlignChoice[] = [
  { label: "Align left", icon: AlignLeft, align: "left" },
  { label: "Align center", icon: AlignCenter, align: "center" },
  { label: "Align right", icon: AlignRight, align: "right" },
  { label: "Justify", icon: AlignJustify, align: "justify" },
];

/** Uses the left-align icon when the selection has no shared alignment. */
export function alignIcon(current: ActiveParagraphAlign): LucideIcon {
  const align = current.kind === "shared" ? current.align : null;
  return ALIGNMENTS.find((choice) => choice.align === align)?.icon ?? AlignLeft;
}

export interface AlignMenuProps {
  state: EditorState;
  run: RunCommand;
  close: () => void;
  takeFocus: boolean;
}

export function AlignMenu({
  state,
  run,
  close,
  takeFocus,
}: AlignMenuProps): ReactElement {
  const current = activeParagraphAlign(state);
  const menu = useRef<HTMLDivElement | null>(null);
  const keys = useMenuKeyboard({ menu, onClose: close, takeFocus });

  return (
    <div
      ref={menu}
      className={editorClassNames.menuList}
      role="menu"
      aria-label="Paragraph alignment"
      {...keys}
    >
      {ALIGNMENTS.map((choice) => (
        <MenuChoice
          key={choice.align}
          label={choice.label}
          icon={choice.icon}
          role="menuitemradio"
          checked={current.kind === "shared" && current.align === choice.align}
          onChoose={() => {
            run(setParagraphAlign(choice.align));
            close();
          }}
        />
      ))}
    </div>
  );
}
