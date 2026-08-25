/** Selects document-defined paragraph styles and clears explicit styles through the default entry. */

import type { ReactElement } from "react";
import type { ParagraphStyleOption } from "../docx/formatting";
import {
  type ActiveParagraphStyle,
  setParagraphStyle,
} from "../editor/commands/paragraphCommands";
import { editorClassNames } from "../styles/classNames";
import type { RunCommand } from "./runCommand";
import { Tooltip } from "./Tooltip";

const LABEL = "Style";

const CLEAR = "default";

const BLANK = "";

/** A style id is any string at all, so it is carried behind a prefix no sentinel value can wear */
const ID_PREFIX = "id:";

const styleValue = (id: string) => `${ID_PREFIX}${id}`;

function defaultName(styles: readonly ParagraphStyleOption[]): string {
  return styles.find((style) => style.isDefault)?.name ?? "Normal";
}

function namedStyles(
  styles: readonly ParagraphStyleOption[]
): ParagraphStyleOption[] {
  return styles.filter(
    (style) => !style.isDefault && style.primary && !style.hidden
  );
}

function selectedValue(
  styles: readonly ParagraphStyleOption[],
  current: ActiveParagraphStyle
): string {
  if (current.kind !== "shared") return BLANK;
  if (current.styleId === null) return CLEAR;
  const named = styles.find((style) => style.id === current.styleId);
  if (!named || named.hidden || (!named.isDefault && !named.primary)) {
    return BLANK;
  }
  return named?.isDefault ? CLEAR : styleValue(current.styleId);
}

export interface ParagraphStyleSelectProps {
  current: ActiveParagraphStyle;
  styles: readonly ParagraphStyleOption[];
  run: RunCommand;
}

export function ParagraphStyleSelect({
  current,
  styles,
  run,
}: ParagraphStyleSelectProps): ReactElement {
  const value = selectedValue(styles, current);
  return (
    <Tooltip label={LABEL}>
      <select
        className={editorClassNames.toolbarSelect}
        aria-label={LABEL}
        value={value}
        disabled={current.kind === "none"}
        onChange={(event) => {
          const chosen = event.target.value;
          if (chosen === BLANK) return;
          run(
            setParagraphStyle(
              chosen === CLEAR ? null : chosen.slice(ID_PREFIX.length)
            )
          );
        }}
      >
        {value === BLANK && <option value={BLANK} />}
        <option value={CLEAR}>{defaultName(styles)}</option>
        {namedStyles(styles).map((style) => (
          <option key={style.id} value={styleValue(style.id)}>
            {style.name}
          </option>
        ))}
      </select>
    </Tooltip>
  );
}
