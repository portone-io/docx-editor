/** Preset and custom colors shared by text, background, border, and cell formatting. */

import { Check } from "lucide-react";
import { type KeyboardEvent, useRef, useState } from "react";
import { editorClassNames } from "../styles/classNames";
import { type ColorRow, DEFAULT_COLORS } from "../styles/colors";
import { useGridKeyboard } from "./useGridKeyboard";

const FALLBACK_CUSTOM_COLOR = "#000000";

function normalizeHexColor(value: string): string | null {
  const digits = value.trim().replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(digits)) {
    return `#${[...digits]
      .map((digit) => `${digit}${digit}`)
      .join("")}`.toUpperCase();
  }
  return /^[0-9a-f]{6}$/i.test(digits) ? `#${digits.toUpperCase()}` : null;
}

function sameColor(a: string | null, b: string): boolean {
  return a !== null && a.toLowerCase() === b.toLowerCase();
}

export interface ColorPickerProps {
  current: string | null;
  colors?: readonly ColorRow[];
  /** Label for the entry that clears the color. */
  clearLabel?: string;
  close: () => void;
  takeFocus: boolean;
  onPick: (hex: string | null) => void;
}

export function ColorPicker({
  current,
  colors = DEFAULT_COLORS,
  clearLabel = "None",
  close,
  takeFocus,
  onPick,
}: ColorPickerProps) {
  const grid = useRef<HTMLDivElement | null>(null);
  const nativePicker = useRef<HTMLInputElement | null>(null);
  const applyButton = useRef<HTMLButtonElement | null>(null);
  const initialCustomColor = normalizeHexColor(current ?? "");
  const [customColor, setCustomColor] = useState(
    initialCustomColor ?? FALLBACK_CUSTOM_COLOR
  );
  const normalizedCustomColor = normalizeHexColor(customColor);
  const keys = useGridKeyboard({
    grid,
    onClose: close,
    start: (cells) =>
      cells.find((cell) => cell.getAttribute("aria-selected") === "true"),
    takeFocus,
  });

  const selectedCell = () =>
    grid.current?.querySelector<HTMLElement>('[tabindex="0"]');

  const onPanelKeyDownCapture = (event: KeyboardEvent<HTMLDivElement>) => {
    if (
      event.key !== "Tab" ||
      event.shiftKey ||
      !(event.target instanceof Element) ||
      !event.target.closest('[role="grid"]')
    ) {
      return;
    }
    nativePicker.current?.focus();
    event.preventDefault();
    event.stopPropagation();
  };

  const onCustomKeyDown = (event: KeyboardEvent<HTMLFormElement>) => {
    if (event.key !== "Tab") return;
    if (event.shiftKey && event.target === nativePicker.current) {
      selectedCell()?.focus({ preventScroll: true });
      event.preventDefault();
      return;
    }
    if (!event.shiftKey && event.target === applyButton.current) {
      close();
      event.preventDefault();
    }
  };

  return (
    <div
      className={editorClassNames.colorPicker}
      onKeyDownCapture={onPanelKeyDownCapture}
    >
      <div
        ref={grid}
        className={editorClassNames.colorGrid}
        role="grid"
        aria-label="Colors"
        {...keys}
      >
        <div className={editorClassNames.colorClearRow} role="row">
          <button
            type="button"
            role="gridcell"
            className={editorClassNames.menuItem}
            aria-selected={current === null}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => onPick(null)}
          >
            {clearLabel}
          </button>
        </div>
        {colors.map((row) => (
          <div key={row[0]} className={editorClassNames.colorRow} role="row">
            {row.map((hex) => (
              <button
                key={hex}
                type="button"
                role="gridcell"
                className={editorClassNames.swatch}
                aria-label={hex}
                aria-selected={sameColor(current, hex)}
                style={{ background: hex }}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => onPick(hex)}
              />
            ))}
          </div>
        ))}
      </div>
      <form
        className={editorClassNames.colorCustom}
        onKeyDown={onCustomKeyDown}
        onSubmit={(event) => {
          event.preventDefault();
          if (normalizedCustomColor) onPick(normalizedCustomColor);
        }}
      >
        <input
          ref={nativePicker}
          type="color"
          className={editorClassNames.colorNativeInput}
          aria-label="Choose custom color"
          value={
            normalizedCustomColor ?? initialCustomColor ?? FALLBACK_CUSTOM_COLOR
          }
          onChange={(event) => {
            const color = normalizeHexColor(event.target.value);
            if (color) setCustomColor(color);
          }}
        />
        <input
          type="text"
          className={editorClassNames.colorHexInput}
          aria-label="Custom color"
          aria-invalid={customColor.length > 0 && !normalizedCustomColor}
          autoComplete="off"
          maxLength={7}
          placeholder="#RRGGBB"
          spellCheck={false}
          value={customColor}
          onChange={(event) => setCustomColor(event.target.value)}
        />
        <button
          ref={applyButton}
          type="submit"
          className={editorClassNames.colorApply}
          aria-label="Apply color"
          disabled={!normalizedCustomColor}
        >
          <Check size={16} aria-hidden="true" />
        </button>
      </form>
    </div>
  );
}
