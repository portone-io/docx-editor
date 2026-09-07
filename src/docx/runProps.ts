/**
 * Changes the character formatting XML (`<w:rPr>`) one item at a time.
 *
 * Turning bold on and off must not lose the other formatting that run wrote down (fonts, shading,
 * even things we do not know about). So instead of building a new fragment we swap out just the one
 * child. It is the same principle `tableFormatting` applies to `w:tcPr`.
 *
 * Which children a job swaps, and what it swaps them for, is the run property table's to say
 * (`formatting/runProperties`); this module holds the surgery around it.
 *
 * The display values produced here are used on screen only. What goes back into the document is
 * always the rPr string.
 */

import { type RunFormat, toRunFormat } from "../model/format";
import {
  type Props,
  parseProps,
  parsePropsXml,
  renderProps,
} from "../ooxml/props";
import {
  type EditableRunKey,
  EMPTY_RUN_PROPS,
  RUN_PROPERTIES,
  type RunEdit,
  readRunFormat,
  runChildEdits,
  withChildEdits,
} from "./formatting";

export type { RunEdit } from "./formatting";
export { matchesRunEdit } from "./formatting";

/** The character formatting that is toggled on and off */
export type RunToggle = "bold" | "italic" | "underline" | "strike";

/** The formatting a run holds. The original XML and the display values derived from that XML form a pair */
export interface RunProps {
  /** The whole `<w:rPr>...</w:rPr>` XML. null for a run with no formatting */
  rPr: string | null;
  format: RunFormat | null;
}

function propsOf(rPr: string | null): Props | null {
  return rPr === null ? EMPTY_RUN_PROPS : parseProps(rPr);
}

/** Derives the display values again from the rPr we operated on. The same rPr always yields the same display values */
export function readRunProps(rPr: string | null): RunFormat | null {
  return rPr === null ? null : readRunFormat(parsePropsXml(rPr));
}

/**
 * The part of the display values that came from the paragraph style rather than from the rPr.
 * Our surgery only touches the rPr, so the values the style gave are left as they are.
 */
function inheritedFormat(current: RunProps): Record<string, unknown> {
  const direct = new Set(Object.keys(readRunProps(current.rPr) ?? {}));
  const inherited: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(current.format ?? {})) {
    if (!direct.has(key)) inherited[key] = value;
  }
  return inherited;
}

function nextProps(rPr: string | null, previous: RunProps): RunProps {
  const inherited = inheritedFormat(previous);
  const format = toRunFormat({ ...inherited, ...(readRunProps(rPr) ?? {}) });
  // A run left with no formatting at all goes back to its initial state, with neither an rPr nor display values
  if (rPr === null && (format === null || Object.keys(format).length === 0)) {
    return { rPr: null, format: null };
  }
  return { rPr, format };
}

/**
 * The rPr and display values after changing one piece of character formatting.
 *
 * `inherited` is everything the layers below the run lay down (`inheritedRunFormat` in
 * `formatting`), which decides whether an off is pinned down or the element simply removed.
 * null for an rPr whose shape could not be made out, or for a value that cannot be written into
 * the document, in which case the original is left untouched.
 */
export function editRunProps<K extends EditableRunKey>(
  current: RunProps,
  inherited: RunFormat,
  edit: RunEdit<K>
): RunProps | null {
  const props = propsOf(current.rPr);
  if (!props) return null;

  const edits = runChildEdits(edit, { rPr: props, inherited });
  if (!edits) return null;

  const rPr = renderProps(withChildEdits(props, edits));
  return nextProps(rPr === "" ? null : rPr, current);
}

/** Whether one of the toggled formats is on in these display values */
export function isRunToggleOn(
  format: RunFormat | null,
  toggle: RunToggle
): boolean {
  return RUN_PROPERTIES[toggle].isOn(format ?? {});
}
