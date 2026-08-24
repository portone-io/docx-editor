/** Table context menu with an optional Unlock action for whole-cell locks. */

import type { Command, EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { Fragment, useCallback, useRef } from "react";
import { unlockSelection } from "../editor/commands/lockCommands";
import {
  closeTableMenu,
  type TableMenuAnchor,
} from "../editor/plugins/tableContextMenu";
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

export interface TableMenuProps {
  view: EditorView;
  state: EditorState;
  anchor: TableMenuAnchor;
  allowLocking?: boolean;
}

export function TableMenu({
  view,
  state,
  anchor,
  allowLocking = false,
}: TableMenuProps) {
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

  const groups = allowLocking ? [...MENU_GROUPS, UNLOCK_GROUP] : MENU_GROUPS;

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
