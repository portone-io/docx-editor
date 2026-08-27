/** Reads the active inline formatting shown by toolbar controls. */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { fontNamesOf } from "../../../docx/formatting";
import {
  type DocumentDefaults,
  NO_FILL,
  type RunFormat,
  toRunFormat,
} from "../../../model/format";
import {
  comparableFontName,
  DEFAULT_FONT_FALLBACKS,
  type FontFallbacks,
} from "../../../styles/fontStack";
import {
  FALLBACK_FONT_SIZE_PT,
  HIGHLIGHT_COLORS,
} from "../../../styles/inlineStyle";
import { documentDefaults } from "../../documentStyles";
import { activePieces, runMarkOf } from "./shared";

/** Null when different values are mixed across the selection (the toolbar shows this as "multiple values") */
function sharedValue<T>(values: (T | null)[]): T | null {
  if (values.length === 0) return null;
  const [first, ...rest] = values;
  return rest.every((value) => value === first) ? first : null;
}

/**
 * The font size at the selected position.
 *
 * "nobody wrote down a size" is kept apart from "several sizes are mixed".
 * Lumping the two together would leave the toolbar unable to tell which is which, and both
 * would show as the default value.
 */
export type ActiveFontSize =
  /** Every selected character has the same size. A size given by a paragraph style lands here too */
  | { kind: "size"; pt: number }
  /** Nobody wrote down a size, so `pt` is the size that actually gets rendered */
  | { kind: "default"; pt: number }
  /** There are several sizes */
  | { kind: "mixed" };

/**
 * The size that actually gets rendered when nobody wrote one down.
 * The document default if there is one, otherwise the on-screen fallback (10pt, same as Word).
 */
function defaultFontSizePt(state: EditorState): number {
  return documentDefaults(state).fontSizePt ?? FALLBACK_FONT_SIZE_PT;
}

export function activeFontSize(state: EditorState): ActiveFontSize {
  const sizes = activePieces(state).map(
    (target) => target.format?.fontSizePt ?? null
  );
  const first = sizes[0] ?? null;
  if (sizes.some((pt) => pt !== first)) return { kind: "mixed" };
  if (first !== null) return { kind: "size", pt: first };
  return { kind: "default", pt: defaultFontSizePt(state) };
}

/**
 * The font at the selected position.
 * The same three branches as size: explicit, unspecified (the name that actually gets rendered),
 * and mixed.
 */
export type ActiveFontFamily =
  /** Every selected character has the same font. A font given by a paragraph style lands here too */
  | { kind: "font"; name: string }
  /** Nobody wrote down a font, so `name` is the font that actually gets rendered */
  | { kind: "default"; name: string }
  /** There are several fonts */
  | { kind: "mixed" };

/** The font name applied to this text. When different names are written per script slot, it is the first one */
function fontNameOf(format: RunFormat | null): string | null {
  return fontNamesOf(format?.fontFamily)[0] ?? null;
}

/**
 * The font name that actually gets rendered when nobody wrote one down.
 * The document default if there is one, otherwise the font the fallback stack provides.
 */
function defaultFontName(
  state: EditorState,
  fontFallbacks: FontFallbacks
): string {
  const defaults = documentDefaults(state);
  return fontNamesOf(defaults.fontFamily)[0] ?? fontFallbacks.defaultFontName;
}

export function activeFontFamily(
  state: EditorState,
  fontFallbacks: FontFallbacks = DEFAULT_FONT_FALLBACKS
): ActiveFontFamily {
  const names = activePieces(state).map((target) => fontNameOf(target.format));
  const first = names[0] ?? null;
  // Two spellings of one name are the same font, so the comparison runs on the comparable form
  const keys = names.map((name) =>
    name === null ? null : comparableFontName(name)
  );
  if (keys.some((key) => key !== keys[0])) return { kind: "mixed" };
  if (first !== null) return { kind: "font", name: first };
  return { kind: "default", name: defaultFontName(state, fontFallbacks) };
}

/**
 * Every font name this document uses. Collects the document default font and the fonts the
 * runs wrote down. The toolbar select appends them after its presets.
 * This walks the whole document, so the caller only calls it when the document changes.
 */
export function documentFontNames(
  doc: PMNode,
  defaults: DocumentDefaults
): string[] {
  const names = new Map<string, string>();
  const collect = (name: string): void => {
    const comparable = comparableFontName(name);
    if (!names.has(comparable)) names.set(comparable, name);
  };

  for (const name of fontNamesOf(defaults.fontFamily)) collect(name);
  doc.descendants((node) => {
    const mark = runMarkOf(node.marks);
    for (const name of fontNamesOf(
      toRunFormat(mark?.attrs.format)?.fontFamily
    )) {
      collect(name);
    }
    return true;
  });
  return Array.from(names.values()).sort((a, b) => a.localeCompare(b));
}

export function activeTextColor(state: EditorState): string | null {
  return sharedValue(
    activePieces(state).map((target) => target.format?.color ?? null)
  );
}

/**
 * The background color actually visible on this text.
 * Word paints the highlight over the shading, so when a highlight is present its color wins.
 * Highlights are written as names, so they are converted to colors for the palette to compare against.
 */
function backgroundOf(format: RunFormat | null): string | null {
  const highlight = format?.highlight;
  if (highlight !== undefined) return HIGHLIGHT_COLORS[highlight] ?? null;
  const background = format?.background ?? null;
  return background === NO_FILL ? null : background;
}

export function activeTextBackground(state: EditorState): string | null {
  return sharedValue(
    activePieces(state).map((target) => backgroundOf(target.format ?? null))
  );
}
