/** Shows the resolved font, leaves mixed selections blank, and clears explicit fonts via Default. */

import type { ReactElement } from "react";
import {
  type ActiveFontFamily,
  setFontFamily,
} from "../editor/commands/formattingCommands";
import { editorClassNames } from "../styles/classNames";
import {
  comparableFontName,
  type FontFallbacks,
  withFontFallback,
} from "../styles/fontStack";
import { DEFAULT_FONTS } from "../styles/fonts";
import type { RunCommand } from "./runCommand";
import { Tooltip } from "./Tooltip";

const LABEL = "Font";

/** A semicolon cannot occur in a font name, so this sentinel cannot collide. */
const CLEAR = ";default";

const MIXED = "";

/**
 * The names to offer, each spelling of a name kept only once and in the form it was first
 * written in. The currently shown name is included so it stays selectable.
 */
function optionNames(
  current: ActiveFontFamily,
  fonts: readonly string[],
  documentFonts: readonly string[]
): string[] {
  const names = [
    ...fonts,
    ...documentFonts,
    ...(current.kind === "mixed" ? [] : [current.name]),
  ];
  const byComparable = new Map<string, string>();
  for (const name of names) {
    const comparable = comparableFontName(name);
    if (!byComparable.has(comparable)) byComparable.set(comparable, name);
  }
  return Array.from(byComparable.values());
}

/**
 * The option the shown name selects, in that option's own spelling.
 * A value matching no option renders the select blank, which is how a mixed selection is
 * signalled, so a name the list holds under another spelling has to resolve to that one.
 */
function selectedValue(options: readonly string[], name: string): string {
  const comparable = comparableFontName(name);
  return (
    options.find((option) => comparableFontName(option) === comparable) ?? name
  );
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
}: FontFamilySelectProps): ReactElement {
  const options = optionNames(current, fonts, documentFonts);
  const value =
    current.kind === "mixed" ? MIXED : selectedValue(options, current.name);

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
        {options.map((name) => (
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
