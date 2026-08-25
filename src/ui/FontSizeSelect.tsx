/** Shows the resolved size, leaves mixed selections blank, and clears explicit sizes via Default. */

import type { ReactElement } from "react";
import {
  type ActiveFontSize,
  setFontSize,
} from "../editor/commands/formattingCommands";
import { editorClassNames } from "../styles/classNames";
import { DEFAULT_FONT_SIZES } from "../styles/fontSizes";
import type { RunCommand } from "./runCommand";
import { Tooltip } from "./Tooltip";

const LABEL = "Font size";

const CLEAR = "default";

const MIXED = "";

export interface FontSizeSelectProps {
  current: ActiveFontSize;
  fontSizes?: readonly number[];
  disabled?: boolean;
  run: RunCommand;
}

export function FontSizeSelect({
  current,
  fontSizes = DEFAULT_FONT_SIZES,
  disabled,
  run,
}: FontSizeSelectProps): ReactElement {
  const value = current.kind === "mixed" ? MIXED : `${current.pt}`;
  // If the document uses a size that is not in the list, show it too so it stays selectable
  const extra =
    current.kind !== "mixed" && !fontSizes.includes(current.pt)
      ? current.pt
      : null;

  return (
    <Tooltip label={LABEL}>
      <select
        className={editorClassNames.toolbarSelect}
        aria-label={LABEL}
        value={value}
        disabled={disabled}
        onChange={(event) => {
          const chosen = event.target.value;
          if (chosen === MIXED) return;
          run(setFontSize(chosen === CLEAR ? null : Number(chosen)));
        }}
      >
        {current.kind === "mixed" && <option value={MIXED} />}
        <option value={CLEAR}>Default</option>
        {extra !== null && <option value={`${extra}`}>{extra}pt</option>}
        {fontSizes.map((pt) => (
          <option key={pt} value={`${pt}`}>
            {pt}pt
          </option>
        ))}
      </select>
    </Tooltip>
  );
}
