/** Table context menu with footnote insertion and an optional Unlock action for whole-cell locks. */

import type { Command, EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { Fragment, type ReactElement, useCallback, useRef } from "react";
import { unlockSelection } from "../editor/commands/lockCommands";
import {
  closeTableMenu,
  type TableMenuAnchor,
} from "../editor/plugins/tableContextMenu";
import type { SurfaceCapabilities } from "../editor/stories/storyView";
import { editorClassNames } from "../styles/classNames";
import {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  mergeCells,
  splitCell,
} from "../table";
import { footnoteItem } from "./footnoteItem";
import { usePanelAtPoint } from "./panelPlacement";
import { commandRunner } from "./runCommand";
import { useDismiss } from "./useDismiss";
import { useMenuKeyboard } from "./useMenuKeyboard";

interface MenuItem {
  label: string;
  command: Command;
}

interface MenuGroup {
  name: string;
  items: readonly MenuItem[];
}

const MENU_GROUPS: readonly MenuGroup[] = [
  {
    name: "rows",
    items: [
      { label: "Insert row above", command: addRowBefore },
      { label: "Insert row below", command: addRowAfter },
      { label: "Delete row", command: deleteRow },
    ],
  },
  {
    name: "columns",
    items: [
      { label: "Insert column left", command: addColumnBefore },
      { label: "Insert column right", command: addColumnAfter },
      { label: "Delete column", command: deleteColumn },
    ],
  },
  {
    name: "cells",
    items: [
      { label: "Merge cells", command: mergeCells },
      { label: "Split cell", command: splitCell },
    ],
  },
  {
    name: "table",
    items: [{ label: "Delete table", command: deleteTable }],
  },
];

const UNLOCK_GROUP: MenuGroup = {
  name: "lock",
  items: [{ label: "Unlock", command: unlockSelection }],
};

/**
 * The note group, which this menu is the only pointer path to for a caret in a cell: a click
 * landing in one with nothing selected is this menu's rather than the text menu's
 * (`editor/plugins/textContextMenu`). The shortcut is left off the row, as every row here leaves
 * its own off.
 */
function noteGroup(takes: SurfaceCapabilities): MenuGroup | null {
  const footnote = footnoteItem(takes);
  return footnote === null
    ? null
    : {
        name: "note",
        items: [{ label: footnote.label, command: footnote.command }],
      };
}

export interface TableMenuProps {
  view: EditorView;
  state: EditorState;
  anchor: TableMenuAnchor;
  /** What the surface holding the caret takes, which the entries that put something in ask */
  takes: SurfaceCapabilities;
  allowLocking?: boolean;
}

export function TableMenu({
  view,
  state,
  anchor,
  takes,
  allowLocking = false,
}: TableMenuProps): ReactElement {
  const box = useRef<HTMLDivElement | null>(null);
  const placement = usePanelAtPoint(box, anchor);
  const run = commandRunner(view);
  const close = useCallback(() => {
    closeTableMenu(view.state, (tr) => view.dispatch(tr));
    // Context menus take focus when opened.
    view.focus();
  }, [view]);
  useDismiss(box, true, close);
  const keys = useMenuKeyboard({ menu: box, onClose: close });

  const choose = (command: Command) => {
    // aria-disabled keeps the row focusable, so enforce it in the handler.
    if (!command(state)) return;
    run(command);
    close();
  };

  const note = noteGroup(takes);
  const groups: readonly MenuGroup[] = [
    ...MENU_GROUPS,
    ...(note === null ? [] : [note]),
    ...(allowLocking ? [UNLOCK_GROUP] : []),
  ];

  return (
    <div
      ref={box}
      className={editorClassNames.menu}
      role="menu"
      aria-label="Table actions"
      {...keys}
      style={{
        left: placement?.left ?? anchor.clientX,
        top: placement?.top ?? anchor.clientY,
        visibility: placement ? undefined : "hidden",
      }}
    >
      {groups.map((group, index) => (
        <Fragment key={group.name}>
          {index > 0 && <hr className={editorClassNames.menuSeparator} />}
          {group.items.map((item) => (
            <button
              key={item.label}
              type="button"
              role="menuitem"
              className={editorClassNames.menuItem}
              aria-disabled={!item.command(state)}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => choose(item.command)}
            >
              {item.label}
            </button>
          ))}
        </Fragment>
      ))}
    </div>
  );
}
