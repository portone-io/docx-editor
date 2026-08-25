/** Editor toolbar with roving focus across its enabled controls. */

import {
  Baseline,
  Bold,
  Highlighter,
  IndentDecrease,
  IndentIncrease,
  Italic,
  Link,
  List,
  ListChevronsUpDown,
  ListOrdered,
  type LucideIcon,
  MessagesSquare,
  Redo2,
  Strikethrough,
  Table,
  Underline,
  Undo2,
} from "lucide-react";
import type { Command, EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { type KeyboardEvent, type ReactElement, useMemo, useRef } from "react";
import {
  activeFontFamily,
  activeFontSize,
  activeTextBackground,
  activeTextColor,
  canFormatText,
  documentFontNames,
  isBoldActive,
  isItalicActive,
  isStrikeActive,
  isUnderlineActive,
  setTextBackground,
  setTextColor,
  toggleBold,
  toggleItalic,
  toggleStrike,
  toggleUnderline,
} from "../editor/commands/formattingCommands";
import { redo, undo } from "../editor/commands/historyCommands";
import {
  canDecreaseIndent,
  canIncreaseIndent,
  decreaseIndent,
  increaseIndent,
} from "../editor/commands/indentCommands";
import {
  activeListKind,
  toggleBulletList,
  toggleNumberedList,
} from "../editor/commands/listCommands";
import {
  activeParagraphAlign,
  activeParagraphStyle,
  canSetParagraphAlign,
} from "../editor/commands/paragraphCommands";
import { canSetLineSpacing } from "../editor/commands/spacingCommands";
import {
  documentDefaults,
  documentParagraphStyles,
} from "../editor/documentStyles";
import { canInsertTable } from "../editor/insertTable";
import { openLinkPanel } from "../editor/plugins/linkPanel";
import type { ListKind } from "../numbering/listTemplate";
import { editorClassNames } from "../styles/classNames";
import type { FontFallbacks } from "../styles/fontStack";
import { AlignMenu, alignIcon } from "./AlignMenu";
import { CellStyleGroup } from "./CellStyleGroup";
import { ColorPicker } from "./ColorPicker";
import { FontFamilySelect } from "./FontFamilySelect";
import { FontSizeSelect } from "./FontSizeSelect";
import { InsertImageButton } from "./InsertImageButton";
import { useLinearWalk } from "./keyboardWalk";
import { ParagraphStyleSelect } from "./ParagraphStyleSelect";
import { Popover } from "./Popover";
import type { DocxEditorPresets } from "./presets";
import { commandRunner, type RunCommand } from "./runCommand";
import { SpacingMenu } from "./SpacingMenu";
import { TableSizePicker } from "./TableSizePicker";
import { ToolbarButton } from "./ToolbarButton";
import { hushedAttribute, useTooltipsHushed } from "./Tooltip";
import { ZoomSelect } from "./ZoomSelect";
import type { DocxEditorZoom } from "./zoom";

interface TextToggleSpec {
  label: string;
  icon: LucideIcon;
  command: Command;
  active: (state: EditorState) => boolean;
}

const TEXT_TOGGLES: readonly TextToggleSpec[] = [
  { label: "Bold", icon: Bold, command: toggleBold, active: isBoldActive },
  {
    label: "Italic",
    icon: Italic,
    command: toggleItalic,
    active: isItalicActive,
  },
  {
    label: "Underline",
    icon: Underline,
    command: toggleUnderline,
    active: isUnderlineActive,
  },
  {
    label: "Strikethrough",
    icon: Strikethrough,
    command: toggleStrike,
    active: isStrikeActive,
  },
];

interface ListToggleSpec {
  label: string;
  icon: LucideIcon;
  kind: ListKind;
  command: Command;
}

const LIST_TOGGLES: readonly ListToggleSpec[] = [
  {
    label: "Numbered list",
    icon: ListOrdered,
    kind: "numbered",
    command: toggleNumberedList,
  },
  {
    label: "Bulleted list",
    icon: List,
    kind: "bullet",
    command: toggleBulletList,
  },
];

interface IndentSpec {
  label: string;
  icon: LucideIcon;
  command: Command;
  enabled: (state: EditorState) => boolean;
}

const INDENTS: readonly IndentSpec[] = [
  {
    label: "Decrease indent",
    icon: IndentDecrease,
    command: decreaseIndent,
    enabled: canDecreaseIndent,
  },
  {
    label: "Increase indent",
    icon: IndentIncrease,
    command: increaseIndent,
    enabled: canIncreaseIndent,
  },
];

function Separator() {
  return <span className={editorClassNames.toolbarSeparator} />;
}

/** Excludes controls inside open popovers, which handle their own keys. */
const TOOLBAR_CONTROLS = `.${editorClassNames.toolbarButton},.${editorClassNames.toolbarSelect}`;

function isLive(control: HTMLElement): boolean {
  return !control.matches(":disabled");
}

function HistoryGroup({ state, run }: { state: EditorState; run: RunCommand }) {
  return (
    <div className={editorClassNames.toolbarGroup}>
      <ToolbarButton
        label="Undo"
        icon={Undo2}
        disabled={!undo(state)}
        onRun={() => run(undo)}
      />
      <ToolbarButton
        label="Redo"
        icon={Redo2}
        disabled={!redo(state)}
        onRun={() => run(redo)}
      />
    </div>
  );
}

function TextToggleGroup({
  state,
  run,
  disabled,
}: {
  state: EditorState;
  run: RunCommand;
  disabled: boolean;
}) {
  return (
    <div className={editorClassNames.toolbarGroup}>
      {TEXT_TOGGLES.map((toggle) => (
        <ToolbarButton
          key={toggle.label}
          label={toggle.label}
          icon={toggle.icon}
          pressed={toggle.active(state)}
          disabled={disabled}
          onRun={() => run(toggle.command)}
        />
      ))}
    </div>
  );
}

function ColorGroup({
  state,
  colors,
  run,
  disabled,
}: {
  state: EditorState;
  colors: DocxEditorPresets["colors"];
  run: RunCommand;
  disabled: boolean;
}) {
  return (
    <div className={editorClassNames.toolbarGroup}>
      <Popover
        label="Text color"
        icon={Baseline}
        indicatorColor={activeTextColor(state) ?? "#000000"}
        panel="dialog"
        disabled={disabled}
      >
        {({ close, takeFocus }) => (
          <ColorPicker
            current={activeTextColor(state)}
            colors={colors}
            close={close}
            takeFocus={takeFocus}
            onPick={(hex) => {
              run(setTextColor(hex));
              close();
            }}
          />
        )}
      </Popover>
      <Popover
        label="Highlight"
        icon={Highlighter}
        indicatorColor={activeTextBackground(state) ?? "transparent"}
        panel="dialog"
        disabled={disabled}
      >
        {({ close, takeFocus }) => (
          <ColorPicker
            current={activeTextBackground(state)}
            colors={colors}
            close={close}
            takeFocus={takeFocus}
            onPick={(hex) => {
              run(setTextBackground(hex));
              close();
            }}
          />
        )}
      </Popover>
    </div>
  );
}

function ParagraphLayoutGroup({
  state,
  lineSpacings,
  run,
}: {
  state: EditorState;
  lineSpacings: DocxEditorPresets["lineSpacings"];
  run: RunCommand;
}) {
  return (
    <div className={editorClassNames.toolbarGroup}>
      <Popover
        label="Alignment"
        icon={alignIcon(activeParagraphAlign(state))}
        panel="menu"
        disabled={!canSetParagraphAlign(state)}
      >
        {({ close, takeFocus }) => (
          <AlignMenu
            state={state}
            run={run}
            close={close}
            takeFocus={takeFocus}
          />
        )}
      </Popover>
      <Popover
        label="Line and paragraph spacing"
        icon={ListChevronsUpDown}
        panel="menu"
        disabled={!canSetLineSpacing(state)}
      >
        {({ close, takeFocus }) => (
          <SpacingMenu
            state={state}
            lineSpacings={lineSpacings}
            run={run}
            close={close}
            takeFocus={takeFocus}
          />
        )}
      </Popover>
    </div>
  );
}

function ListGroup({ state, run }: { state: EditorState; run: RunCommand }) {
  const kind = activeListKind(state);
  return (
    <div className={editorClassNames.toolbarGroup}>
      {LIST_TOGGLES.map((toggle) => (
        <ToolbarButton
          key={toggle.kind}
          label={toggle.label}
          icon={toggle.icon}
          pressed={kind === toggle.kind}
          // A document without numbering.xml cannot start a new list
          disabled={!toggle.command(state)}
          onRun={() => run(toggle.command)}
        />
      ))}
    </div>
  );
}

function IndentGroup({ state, run }: { state: EditorState; run: RunCommand }) {
  return (
    <div className={editorClassNames.toolbarGroup}>
      {INDENTS.map((indent) => (
        <ToolbarButton
          key={indent.label}
          label={indent.label}
          icon={indent.icon}
          disabled={!indent.enabled(state)}
          onRun={() => run(indent.command)}
        />
      ))}
    </div>
  );
}

export interface ToolbarProps {
  view: EditorView;
  state: EditorState;
  fontFallbacks?: FontFallbacks;
  /** Optional values offered by toolbar pickers. */
  presets?: DocxEditorPresets;
  onToggleComments?: () => void;
  commentsOpen?: boolean;
  zoom: DocxEditorZoom;
  onZoomChange: (zoom: DocxEditorZoom) => void;
}

export function Toolbar({
  view,
  state,
  fontFallbacks,
  presets,
  onToggleComments,
  commentsOpen = false,
  zoom,
  onZoomChange,
}: ToolbarProps): ReactElement {
  const run = commandRunner(view);
  const bar = useRef<HTMLDivElement | null>(null);
  const keys = useLinearWalk({
    container: bar,
    selector: TOOLBAR_CONTROLS,
    orientation: "horizontal",
    navigable: isLive,
  });
  const hushed = useTooltipsHushed();
  const doc = state.doc;
  const defaults = documentDefaults(state);
  const formattable = canFormatText(state);
  // Font discovery walks the document, so selection-only updates reuse the result.
  const documentFonts = useMemo(
    () => documentFontNames(doc, defaults),
    [doc, defaults]
  );

  // Claim horizontal navigation from selects but leave Up/Down to the native control.
  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const { target } = event;
    if (target instanceof HTMLElement && target.matches(TOOLBAR_CONTROLS)) {
      keys.walk(event);
    }
  };

  return (
    <div
      ref={bar}
      className={editorClassNames.toolbar}
      role="toolbar"
      aria-label="Editor toolbar"
      onKeyDown={onKeyDown}
      {...hushedAttribute(hushed)}
    >
      <HistoryGroup state={state} run={run} />
      <ZoomSelect zoom={zoom} onZoomChange={onZoomChange} />
      <Separator />
      <div className={editorClassNames.toolbarGroup}>
        <ParagraphStyleSelect
          current={activeParagraphStyle(state)}
          styles={documentParagraphStyles(state)}
          run={run}
        />
        <FontFamilySelect
          current={activeFontFamily(state, fontFallbacks)}
          documentFonts={documentFonts}
          fonts={presets?.fonts}
          fontFallbacks={fontFallbacks}
          disabled={!formattable}
          run={run}
        />
        <FontSizeSelect
          current={activeFontSize(state)}
          fontSizes={presets?.fontSizes}
          disabled={!formattable}
          run={run}
        />
      </div>
      <Separator />
      <TextToggleGroup state={state} run={run} disabled={!formattable} />
      <ColorGroup
        state={state}
        colors={presets?.colors}
        run={run}
        disabled={!formattable}
      />
      <Separator />
      <ParagraphLayoutGroup
        state={state}
        lineSpacings={presets?.lineSpacings}
        run={run}
      />
      <Separator />
      <ListGroup state={state} run={run} />
      <IndentGroup state={state} run={run} />
      <Separator />
      <CellStyleGroup
        state={state}
        cellBorders={presets?.cellBorders}
        colors={presets?.colors}
        run={run}
      />
      <Separator />
      <div className={editorClassNames.toolbarGroup}>
        <Popover
          label="Insert table"
          icon={Table}
          panel="dialog"
          disabled={!canInsertTable(state)}
        >
          {({ close, takeFocus }) => (
            <TableSizePicker run={run} close={close} takeFocus={takeFocus} />
          )}
        </Popover>
        <InsertImageButton view={view} state={state} />
        {/* Cmd+K shares this panel, so it anchors to the selection rather than the button. */}
        <ToolbarButton
          label="Link"
          icon={Link}
          disabled={!openLinkPanel(state)}
          onRun={() => run(openLinkPanel)}
        />
        <ToolbarButton
          label={commentsOpen ? "Hide comments" : "Show comments"}
          icon={MessagesSquare}
          pressed={commentsOpen}
          disabled={!onToggleComments}
          onRun={() => onToggleComments?.()}
        />
      </div>
    </div>
  );
}
