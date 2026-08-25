/** Custom text context menu with clipboard, comment, and optional locking actions. */

import {
  ClipboardPaste,
  Copy,
  Lock,
  LockOpen,
  type LucideIcon,
  MessageSquarePlus,
  Scissors,
  Trash2,
} from "lucide-react";
import { deleteSelection } from "prosemirror-commands";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { Fragment, type ReactElement, useCallback, useRef } from "react";
import { canAddComment } from "../editor/commands/commentCommands";
import {
  lockSelection,
  type SelectionLock,
  selectionLock,
  selectionTouchesLocked,
  unlockSelection,
} from "../editor/commands/lockCommands";
import { insertClipboardData } from "../editor/externalClipboard";
import {
  closeTextMenu,
  type TextMenuAnchor,
} from "../editor/plugins/textContextMenu";
import { editorClassNames } from "../styles/classNames";
import { usePanelAtPoint } from "./panelPlacement";
import { commandRunner, type RunCommand } from "./runCommand";
import { ICON_SIZE } from "./ToolbarButton";
import { useDismiss } from "./useDismiss";
import { useMenuKeyboard } from "./useMenuKeyboard";

const MOD =
  typeof navigator !== "undefined" && navigator.platform.includes("Mac")
    ? "⌘"
    : "Ctrl+";

/**
 * Cut and copy are handed to the browser from inside the click that asked for them, which is the
 * only moment a page is allowed to write to the clipboard. What lands there is then whatever
 * ProseMirror itself puts on the clipboard, so this menu keeps no copying path of its own.
 */
function clipboardCommand(view: EditorView, command: "cut" | "copy"): void {
  view.focus();
  document.execCommand(command);
}

interface ClipboardContent {
  html: string;
  text: string;
}

async function clipboardContent(): Promise<ClipboardContent> {
  const clipboard = navigator.clipboard;
  if (!clipboard) return { html: "", text: "" };
  try {
    if (clipboard.read) {
      const items = await clipboard.read();
      let html = "";
      let text = "";
      for (const item of items) {
        if (!html && item.types.includes("text/html")) {
          html = await (await item.getType("text/html")).text();
        }
        if (!text && item.types.includes("text/plain")) {
          text = await (await item.getType("text/plain")).text();
        }
      }
      if (html || text) return { html, text };
    }
  } catch {
    // The plain-text API may still be permitted when the richer read is not.
  }
  return {
    html: "",
    text: await clipboard.readText().catch(() => ""),
  };
}

async function pasteFromClipboard(view: EditorView): Promise<void> {
  const content = await clipboardContent();
  if (!content.html && !content.text) return;
  view.focus();
  insertClipboardData(view, content);
}

interface MenuItem {
  label: string;
  icon: LucideIcon;
  /** The keyboard shortcut for the same action, written on the right of the row */
  hint?: string;
  enabled: boolean;
  run: () => void;
}

/** Mixed locked and unlocked content offers Unlock first. */
function lockItem(lock: SelectionLock, run: RunCommand): MenuItem {
  const locking = {
    label: "Lock",
    icon: Lock,
    run: () => run(lockSelection),
  };
  switch (lock) {
    case "none":
      return { ...locking, enabled: false };
    case "lockable":
      return { ...locking, enabled: true };
    case "locked":
    case "mixed":
      return {
        label: "Unlock",
        icon: LockOpen,
        enabled: true,
        run: () => run(unlockSelection),
      };
  }
}

export interface TextMenuProps {
  view: EditorView;
  state: EditorState;
  anchor: TextMenuAnchor;
  /** Whether the entries that lock and unlock a stretch of text are offered */
  allowLocking?: boolean;
  /** Opens the built-in composer for the selected text. */
  onAddComment?: () => void;
}

export function TextMenu({
  view,
  state,
  anchor,
  allowLocking = false,
  onAddComment,
}: TextMenuProps): ReactElement {
  const box = useRef<HTMLDivElement | null>(null);
  const placement = usePanelAtPoint(box, anchor);
  const run = commandRunner(view);
  const close = useCallback(() => {
    closeTextMenu(view.state, (tr) => view.dispatch(tr));
    // Context menus take focus when opened.
    view.focus();
  }, [view]);
  useDismiss(box, true, close);
  const keys = useMenuKeyboard({ menu: box, onClose: close });

  const selected = !state.selection.empty;
  const shut = selectionTouchesLocked(state);
  const groups: MenuItem[][] = [
    [
      {
        label: "Cut",
        icon: Scissors,
        hint: `${MOD}X`,
        enabled: selected && !shut,
        run: () => clipboardCommand(view, "cut"),
      },
      {
        label: "Copy",
        icon: Copy,
        hint: `${MOD}C`,
        enabled: selected,
        run: () => clipboardCommand(view, "copy"),
      },
      {
        label: "Paste",
        icon: ClipboardPaste,
        hint: `${MOD}V`,
        enabled: !shut,
        run: () => {
          void pasteFromClipboard(view);
        },
      },
      {
        label: "Delete",
        icon: Trash2,
        enabled: selected && !shut,
        run: () => run(deleteSelection),
      },
    ],
  ];
  groups.push([
    {
      label: "Add comment",
      icon: MessageSquarePlus,
      enabled: !!onAddComment && canAddComment(state),
      run: () => onAddComment?.(),
    },
  ]);
  if (allowLocking) groups.push([lockItem(selectionLock(state), run)]);

  const choose = (item: MenuItem) => {
    // aria-disabled keeps the row focusable, so enforce it in the handler.
    if (!item.enabled) return;
    item.run();
    close();
  };

  return (
    <div
      ref={box}
      className={editorClassNames.menu}
      role="menu"
      aria-label="Text actions"
      {...keys}
      style={{
        left: placement?.left ?? anchor.clientX,
        top: placement?.top ?? anchor.clientY,
        visibility: placement ? undefined : "hidden",
      }}
    >
      {groups.map((group, index) => (
        <Fragment key={group[0].label}>
          {index > 0 && <hr className={editorClassNames.menuSeparator} />}
          {group.map((item) => {
            const Icon = item.icon;
            return (
              <button
                key={item.label}
                type="button"
                role="menuitem"
                className={editorClassNames.menuItem}
                aria-disabled={!item.enabled}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => choose(item)}
              >
                <Icon size={ICON_SIZE} aria-hidden="true" />
                {item.label}
                {item.hint && (
                  <span className={editorClassNames.menuHint}>{item.hint}</span>
                )}
              </button>
            );
          })}
        </Fragment>
      ))}
    </div>
  );
}
