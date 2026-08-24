/** Shows the resolved font, leaves mixed selections blank, and clears explicit fonts via Default. */

import {
  type ActiveFontFamily,
  setFontFamily,
} from "../editor/commands/formattingCommands";
import { editorClassNames } from "../styles/classNames";
import { type FontFallbacks, withFontFallback } from "../styles/fontStack";
import { DEFAULT_FONTS } from "../styles/fonts";
import type { RunCommand } from "./runCommand";
import { Tooltip } from "./Tooltip";

const LABEL = "Font";

/** A semicolon cannot occur in a font name, so this sentinel cannot collide. */
const CLEAR = ";default";

const MIXED = "";

/** If the currently shown name is not in the list, include it so it stays selectable */
function optionNames(
  current: ActiveFontFamily,
  fonts: readonly string[],
  documentFonts: readonly string[]
): string[] {
  const names = [...fonts, ...documentFonts];
  if (current.kind !== "mixed" && !names.includes(current.name)) {
    names.push(current.name);
  }
  return Array.from(new Set(names));
}

export interface FontFamilySelectProps {
  current: ActiveFontFamily;
  documentFonts: readonly string[];
  fonts?: readonly string[];
  fontFallbacks?: FontFallbacks;
  disabled?: boolean;
  run: RunCommand;
}

export function FontFamilySelect({
  current,
  documentFonts,
  fonts = DEFAULT_FONTS,
  fontFallbacks,
  disabled,
  run,
}: FontFamilySelectProps) {
  const value = current.kind === "mixed" ? MIXED : current.name;

  return (
    <Tooltip label={LABEL}>
      <select
        className={`${editorClassNames.toolbarSelect} ${editorClassNames.fontFamilySelect}`}
        aria-label={LABEL}
        value={value}
        disabled={disabled}
        onChange={(event) => {
          const chosen = event.target.value;
          if (chosen === MIXED) return;
          run(setFontFamily(chosen === CLEAR ? null : chosen));
        }}
      >
        {current.kind === "mixed" && <option value={MIXED} />}
        <option value={CLEAR}>Default</option>
        {optionNames(current, fonts, documentFonts).map((name) => (
          <option
            key={name}
            value={name}
            style={{ fontFamily: withFontFallback(`"${name}"`, fontFallbacks) }}
          >
            {name}
          </option>
        ))}
      </select>
    </Tooltip>
  );
}
