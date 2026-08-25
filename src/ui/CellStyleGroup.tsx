/** Toolbar controls for formatting the current table cell selection. */

import {
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  Grid2x2,
  type LucideIcon,
  PaintBucket,
  Palette,
  Square,
  SquareDashed,
} from "lucide-react";
import type { EditorState } from "prosemirror-state";
import { type ReactElement, useRef } from "react";
import type { CellVerticalAlign } from "../model/format";
import { editorClassNames } from "../styles/classNames";
import type { ColorRow } from "../styles/colors";
import {
  activeCellBackground,
  activeCellBorderColor,
  activeCellVerticalAlign,
  type CellBorderPreset,
  canSetCellBorderColor,
  canSetCellFormatting,
  setCellBackground,
  setCellBorderColor,
  setCellBorders,
  setCellVerticalAlign,
} from "../table";
import {
  type CellBorderOption,
  DEFAULT_CELL_BORDERS,
} from "../table/cellBorders";
import { ColorPicker } from "./ColorPicker";
import { MenuChoice } from "./MenuChoice";
import { Popover } from "./Popover";
import type { RunCommand } from "./runCommand";
import { ICON_SIZE } from "./ToolbarButton";
import { useMenuKeyboard } from "./useMenuKeyboard";

/** The picture drawn on each preset's row. It is chrome, so it stays out of the exported list */
const BORDER_ICONS: Readonly<Record<CellBorderPreset, LucideIcon>> = {
  all: Grid2x2,
  outer: Square,
  none: SquareDashed,
};

const ALIGNMENTS: readonly {
  value: CellVerticalAlign;
  label: string;
  icon: LucideIcon;
}[] = [
  { value: "top", label: "Align cell top", icon: AlignVerticalJustifyStart },
  {
    value: "center",
    label: "Align cell center",
    icon: AlignVerticalJustifyCenter,
  },
  {
    value: "bottom",
    label: "Align cell bottom",
    icon: AlignVerticalJustifyEnd,
  },
];

function CellVerticalAlignMenu({
  state,
  run,
  close,
  takeFocus,
}: {
  state: EditorState;
  run: RunCommand;
  close: () => void;
  takeFocus: boolean;
}) {
  const current = activeCellVerticalAlign(state);
  const menu = useRef<HTMLDivElement | null>(null);
  const keys = useMenuKeyboard({ menu, onClose: close, takeFocus });

  return (
    <div
      ref={menu}
      className={editorClassNames.menuList}
      role="menu"
      aria-label="Cell vertical alignment"
      {...keys}
    >
      {ALIGNMENTS.map((option) => (
        <MenuChoice
          key={option.value}
          label={option.label}
          icon={option.icon}
          role="menuitemradio"
          checked={current === option.value}
          onChoose={() => {
            run(setCellVerticalAlign(option.value));
            close();
          }}
        />
      ))}
    </div>
  );
}

function cellVerticalAlignIcon(
  current: ReturnType<typeof activeCellVerticalAlign>
): LucideIcon {
  return (
    ALIGNMENTS.find((option) => option.value === current)?.icon ??
    AlignVerticalJustifyStart
  );
}

function BorderPresetMenu({
  state,
  cellBorders,
  run,
  close,
  takeFocus,
}: {
  state: EditorState;
  cellBorders: readonly CellBorderOption[];
  run: RunCommand;
  close: () => void;
  takeFocus: boolean;
}) {
  const menu = useRef<HTMLDivElement | null>(null);
  const keys = useMenuKeyboard({ menu, onClose: close, takeFocus });

  return (
    <div
      ref={menu}
      className={editorClassNames.menuList}
      role="menu"
      aria-label="Cell borders"
      {...keys}
    >
      {cellBorders.map((option) => {
        const command = setCellBorders(option.preset);
        const Icon = BORDER_ICONS[option.preset];
        return (
          <button
            key={option.preset}
            type="button"
            role="menuitem"
            className={editorClassNames.menuItem}
            // Cells already drawn the way the preset wants have nothing to change
            aria-disabled={!command(state)}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => {
              // A row drawn unclickable is still clickable as far as the browser is concerned
              if (!command(state)) return;
              run(command);
              close();
            }}
          >
            <span className={editorClassNames.menuCheck} />
            <Icon size={ICON_SIZE} aria-hidden="true" />
            {option.label}
          </button>
        );
      })}
    </div>
  );
}

export interface CellStyleGroupProps {
  state: EditorState;
  /** The border choices the menu offers. The built-in list when none is given */
  cellBorders?: readonly CellBorderOption[];
  /** The swatches both color buttons offer. The built-in palette when none is given */
  colors?: readonly ColorRow[];
  run: RunCommand;
}

export function CellStyleGroup({
  state,
  cellBorders = DEFAULT_CELL_BORDERS,
  colors,
  run,
}: CellStyleGroupProps): ReactElement {
  const activeAlignment = activeCellVerticalAlign(state);
  const layoutEditable = canSetCellFormatting(state);

  return (
    <div className={editorClassNames.toolbarGroup}>
      <Popover
        label="Cell vertical alignment"
        icon={cellVerticalAlignIcon(activeAlignment)}
        panel="menu"
        disabled={!layoutEditable}
      >
        {({ close, takeFocus }) => (
          <CellVerticalAlignMenu
            state={state}
            run={run}
            close={close}
            takeFocus={takeFocus}
          />
        )}
      </Popover>
      <Popover
        label="Cell borders"
        icon={Grid2x2}
        panel="menu"
        disabled={!layoutEditable}
      >
        {({ close, takeFocus }) => (
          <BorderPresetMenu
            state={state}
            cellBorders={cellBorders}
            run={run}
            close={close}
            takeFocus={takeFocus}
          />
        )}
      </Popover>
      <Popover
        label="Border color"
        icon={Palette}
        indicatorColor={activeCellBorderColor(state) ?? "#000000"}
        panel="dialog"
        disabled={!layoutEditable || !canSetCellBorderColor(state)}
      >
        {({ close, takeFocus }) => (
          <ColorPicker
            current={activeCellBorderColor(state)}
            colors={colors}
            clearLabel="Automatic"
            close={close}
            takeFocus={takeFocus}
            onPick={(hex) => {
              run(setCellBorderColor(hex));
              close();
            }}
          />
        )}
      </Popover>
      <Popover
        label="Cell fill"
        icon={PaintBucket}
        indicatorColor={activeCellBackground(state) ?? "transparent"}
        panel="dialog"
        disabled={!layoutEditable}
      >
        {({ close, takeFocus }) => (
          <ColorPicker
            current={activeCellBackground(state)}
            colors={colors}
            clearLabel="No fill"
            close={close}
            takeFocus={takeFocus}
            onPick={(hex) => {
              run(setCellBackground(hex));
              close();
            }}
          />
        )}
      </Popover>
    </div>
  );
}
