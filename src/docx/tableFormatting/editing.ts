/**
 * Applies focused edits to a cell's preserved formatting XML. The fragment is never rebuilt:
 * only the child and attributes named by the edit are changed, and unsupported formatting stays
 * untouched.
 */

import type {
  CellFormat,
  CellMargins,
  CellVerticalAlign,
  RowFormat,
} from "../../model/format";
import { elementXml, wAttrValue, type XmlAttr } from "../../ooxml/element";
import { wName } from "../../ooxml/names";
import { setAttr } from "../../ooxml/precedence";
import {
  type ChildElement,
  childElement,
  editChild,
  type Props,
  parseProps,
  parsePropsXml,
  propsChild,
  renderProps,
  setChild,
  withAttrs,
} from "../../ooxml/props";
import { normalizeHex } from "../../ooxml/units";
import {
  type CellBorderDefaults,
  NO_BORDER_DEFAULTS,
  NO_CELL_MARGINS,
  readCellProps,
  readRowFormat,
} from "./reading";

export interface CellProps {
  /** The whole `<w:tcPr>...</w:tcPr>` XML. null for a cell with no formatting */
  tcPr: string | null;
  format: CellFormat | null;
}

export type CellSide = "top" | "bottom" | "left" | "right";

export const ALL_CELL_SIDES: readonly CellSide[] = [
  "top",
  "bottom",
  "left",
  "right",
];

/** The line a border preset writes. `none` states that this side draws nothing */
export type CellBorderLine = "single" | "none";

/** A job that changes one piece of cell formatting */
export type CellFormatEdit =
  /** The background fill (`w:shd`). A null color takes the fill away */
  | { kind: "background"; hex: string | null }
  /** The line on the named sides (`w:tcBorders`). The other sides are left alone */
  | { kind: "borders"; line: CellBorderLine; sides: readonly CellSide[] }
  /** The color of the sides that already draw a line. A null color resets them to `auto` */
  | { kind: "borderColor"; hex: string | null }
  /** The vertical placement of content within the cell (`w:vAlign`). */
  | { kind: "verticalAlign"; align: CellVerticalAlign }
  /** The directly formatted cell margins, in points. Omitted sides remain unchanged. */
  | {
      kind: "padding";
      values: Partial<Record<CellSide, number>>;
    };

/** The spellings one side can be written under. A side that is not there yet is written as the first */
const SIDE_NAMES: Record<CellSide, readonly string[]> = {
  top: ["top"],
  bottom: ["bottom"],
  left: ["left", "start"],
  right: ["right", "end"],
};

/** The thickness a preset writes, in eighths of a point (4 = 0.5pt, the line Word draws by default) */
const PRESET_BORDER_EIGHTHS = 4;

const EMPTY_TC_PR: Props = { tag: "w:tcPr", attrs: null, children: [] };

/** Whether this border side draws a line. A missing side, `nil`, and `none` all draw nothing */
function drawsLine(side: ChildElement): boolean {
  if (side.tag === null) return false;
  const val = wAttrValue(side.attrs, "val");
  return val !== null && val !== "nil" && val !== "none";
}

/** What is to become of the sides of `w:tcBorders` */
type BordersEdit =
  | { kind: "line"; line: CellBorderLine; sides: readonly CellSide[] }
  /** Only the sides that already draw a line are recolored. A null color means `auto` */
  | { kind: "color"; hex: string | null };

/** The attributes one side is to be changed to. null leaves that side exactly as it is */
function sideAttrs(
  name: string,
  current: ChildElement,
  edit: BordersEdit,
  fallback: string | null = null
): readonly XmlAttr[] | null {
  if (edit.kind === "color") {
    if (!drawsLine(current)) {
      if (current.tag !== null || !fallback || fallback === "none") return null;
      return inheritedBorderAttrs(fallback, edit.hex);
    }
    return setAttr(current.attrs, name, "color", edit.hex ?? "auto");
  }
  if (edit.line === "none") {
    // The thickness and color stay behind, so a line switched off comes back as it was
    return current.tag !== null
      ? setAttr(current.attrs, name, "val", "none")
      : [[wName("val"), "none"]];
  }
  if (current.tag === null) {
    return [
      [wName("val"), "single"],
      [wName("sz"), `${PRESET_BORDER_EIGHTHS}`],
      [wName("space"), "0"],
      [wName("color"), "auto"],
    ];
  }
  // The color is not ours to decide here, so a themed or explicit color survives the preset
  return setAttr(
    setAttr(current.attrs, name, "val", "single"),
    name,
    "sz",
    `${PRESET_BORDER_EIGHTHS}`
  );
}

const BORDER_VAL_BY_CSS_STYLE: Readonly<Record<string, string>> = {
  solid: "single",
  double: "double",
  dashed: "dashed",
  dotted: "dotted",
};

/** Materializes one visible inherited line as a direct cell border with a new color. */
function inheritedBorderAttrs(
  border: string,
  hex: string | null
): readonly XmlAttr[] | null {
  const matched = border.match(
    /^(\d+(?:\.\d+)?)pt (solid|double|dashed|dotted) (#[0-9a-f]{6})$/i
  );
  if (!matched) return null;
  const [, widthText, style, currentColor] = matched;
  if (hex !== null && normalizeHex(currentColor) === hex) return null;
  const eighths = Math.round(Number.parseFloat(widthText) * 8);
  if (!Number.isSafeInteger(eighths) || eighths <= 0) return null;
  return [
    [wName("val"), BORDER_VAL_BY_CSS_STYLE[style.toLowerCase()]],
    [wName("sz"), `${eighths}`],
    [wName("space"), "0"],
    [wName("color"), hex ?? "auto"],
  ];
}

/** Which spelling of a side to write, and the other spelling to drop so Word is not left with both */
function sideNames(
  borders: Props,
  side: CellSide
): { write: string; drop: string | null } {
  const [primary, alternate] = SIDE_NAMES[side];
  const usesAlternate =
    alternate !== undefined &&
    propsChild(borders.children, primary) === undefined &&
    propsChild(borders.children, alternate) !== undefined;
  if (usesAlternate) return { write: alternate, drop: null };
  return { write: primary, drop: alternate ?? null };
}

/** A leaf element carrying nothing but the attributes it is given */
function leafElement(tag: string, attrs: readonly XmlAttr[]): Props {
  return withAttrs({ tag, attrs: null, children: [] }, attrs);
}

/** The fragment one child of these properties holds, empty when the child is not there */
function containerOf(props: Props, name: string): Props | null {
  const current = propsChild(props.children, name)?.xml;
  if (current === undefined) {
    return { tag: wName(name), attrs: null, children: [] };
  }
  return parseProps(current);
}

/**
 * These cell properties with the sides the job names rewritten inside `<w:tcBorders>`.
 *
 * null if a fragment's shape could not be made out, in which case the caller leaves the cell
 * untouched. A `w:tcBorders` left with no side at all goes away with the last of them.
 */
function editedBorders(
  props: Props,
  edit: BordersEdit,
  defaults: CellBorderDefaults
): Props | null {
  const borders = containerOf(props, "tcBorders");
  if (!borders) return null;

  const sides = edit.kind === "line" ? edit.sides : ALL_CELL_SIDES;
  return sides.reduce<Props | null>((kept, side) => {
    if (kept === null) return null;
    const { write, drop } = sideNames(borders, side);
    const currentSide = childElement(borders, write);
    if (!currentSide) return null;
    const attrs = sideAttrs(write, currentSide, edit, defaults[side]);
    if (!attrs) return kept;
    // The original tag keeps its prefix, so a document not using `w:` is written back as it was
    const written = editChild(kept, ["tcBorders", write], () =>
      leafElement(currentSide.tag ?? wName(write), attrs)
    );
    if (written === null || drop === null) return written;
    return editChild(written, ["tcBorders", drop], () => null);
  }, props);
}

/** Whether the shading paints a pattern rather than a plain fill */
function hasPattern(shd: ChildElement): boolean {
  const val = wAttrValue(shd.attrs, "val");
  return val !== null && val !== "clear" && val !== "nil";
}

/**
 * The `<w:shd>` fragment with the fill written in. null removes the fragment.
 *
 * Only the fill is ours to decide: a pattern and the color it is drawn in stay as they were, and a
 * shading that recorded nothing but a fill goes away entirely once the fill is taken off.
 * A fill needs something to paint it with, so a shading that painted nothing (`nil`, or no `w:val`
 * at all) is moved to `clear`.
 */
function editedShading(
  current: ChildElement,
  fill: string | null
): string | null {
  const tag = current.tag ?? wName("shd");
  if (fill === null) {
    if (!hasPattern(current)) return null;
    return elementXml(tag, setAttr(current.attrs, "shd", "fill", "auto"));
  }
  if (current.tag === null) {
    return elementXml(tag, [
      [wName("val"), "clear"],
      [wName("color"), "auto"],
      [wName("fill"), fill],
    ]);
  }
  const painting = hasPattern(current)
    ? current.attrs
    : setAttr(current.attrs, "shd", "val", "clear");
  return elementXml(tag, setAttr(painting, "shd", "fill", fill));
}

const MARGIN_SIDE_NAMES: Readonly<Record<CellSide, readonly string[]>> = {
  top: ["top"],
  right: ["end", "right"],
  bottom: ["bottom"],
  left: ["start", "left"],
};

/** These cell properties with the sides the job names rewritten inside `<w:tcMar>` */
function editedPadding(
  props: Props,
  values: Partial<Record<CellSide, number>>
): Props | null {
  const margins = containerOf(props, "tcMar");
  if (!margins) return null;
  let edited: Props | null = props;
  let wrote = false;
  for (const side of ALL_CELL_SIDES) {
    const points = values[side];
    if (points === undefined) continue;
    const twips = Math.round(points * 20);
    if (
      !Number.isFinite(points) ||
      points < 0 ||
      !Number.isSafeInteger(twips)
    ) {
      return null;
    }
    // The spelling the document already used for this side, else the newer one of the two
    const name =
      MARGIN_SIDE_NAMES[side].find(
        (spelling) => propsChild(margins.children, spelling) !== undefined
      ) ?? side;
    const existing = childElement(margins, name);
    if (existing === null || edited === null) return null;
    const attrs = setAttr(
      setAttr(existing.attrs, name, "w", `${twips}`),
      name,
      "type",
      "dxa"
    );
    edited = editChild(edited, ["tcMar", name], () =>
      leafElement(existing.tag ?? wName(side), attrs)
    );
    wrote = true;
  }
  return wrote ? edited : null;
}

/** These cell properties after one job. null for a value that cannot be written down */
function editedProps(
  edit: CellFormatEdit,
  props: Props,
  defaults: CellBorderDefaults
): Props | null {
  switch (edit.kind) {
    case "background": {
      const fill = edit.hex === null ? null : normalizeHex(edit.hex);
      if (edit.hex !== null && fill === null) return null;
      const shd = childElement(props, "shd");
      if (!shd) return null;
      return setChild(props, "shd", editedShading(shd, fill));
    }
    case "borders":
      return editedBorders(
        props,
        { kind: "line", line: edit.line, sides: edit.sides },
        defaults
      );
    case "borderColor": {
      const hex = edit.hex === null ? null : normalizeHex(edit.hex);
      if (edit.hex !== null && hex === null) return null;
      return editedBorders(props, { kind: "color", hex }, defaults);
    }
    case "verticalAlign": {
      const current = childElement(props, "vAlign");
      if (!current) return null;
      return setChild(
        props,
        "vAlign",
        elementXml(
          current.tag ?? wName("vAlign"),
          setAttr(current.attrs, "vAlign", "val", edit.align)
        )
      );
    }
    case "padding":
      return editedPadding(props, edit.values);
  }
}

/**
 * The cell formatting XML and display values after one piece of formatting is changed.
 *
 * null when there is nothing to do: a fragment whose shape could not be made out, a value that
 * cannot be written into the document, or a cell already in the state the job wants. In every one
 * of those cases the cell keeps its original XML.
 */
export function editCellProps(
  tcPr: string | null,
  edit: CellFormatEdit,
  defaults: CellBorderDefaults = NO_BORDER_DEFAULTS,
  margins: CellMargins = NO_CELL_MARGINS
): CellProps | null {
  const props = tcPr === null ? EMPTY_TC_PR : parseProps(tcPr);
  if (!props) return null;

  const edited = editedProps(edit, props, defaults);
  if (!edited) return null;

  const rendered = renderProps(edited);
  const next = rendered === "" ? null : rendered;
  if (next === tcPr) return null;
  return { tcPr: next, format: readCellProps(next, defaults, margins) };
}

/**
 * Whether the cell's own formatting draws a line on any side.
 * A border color has nothing to act on when it does not, because the lines on screen are then the
 * table's own, which belong to the table rather than to this cell.
 */
export function drawsOwnCellBorder(tcPr: string | null): boolean {
  const props = tcPr === null ? null : parseProps(tcPr);
  const borders = props && containerOf(props, "tcBorders");
  if (!borders) return false;
  return ALL_CELL_SIDES.some((side) =>
    SIDE_NAMES[side].some((name) => {
      const element = childElement(borders, name);
      return element !== null && drawsLine(element);
    })
  );
}

export interface RowProps {
  trPr: string;
  format: RowFormat;
}

/** Writes a row height in points while retaining unrelated `w:trPr` content. */
export function editRowHeight(
  trPr: string | null,
  heightPt: number
): RowProps | null {
  const twips = Math.round(heightPt * 20);
  if (
    !Number.isFinite(heightPt) ||
    heightPt <= 0 ||
    !Number.isSafeInteger(twips)
  ) {
    return null;
  }
  const props =
    trPr === null
      ? { tag: wName("trPr"), attrs: null, children: [] }
      : parseProps(trPr);
  if (!props) return null;
  const current = childElement(props, "trHeight");
  if (!current) return null;
  const writtenRule = wAttrValue(current.attrs, "hRule");
  const attrs = setAttr(current.attrs, "trHeight", "val", `${twips}`);
  const nextAttrs = setAttr(
    attrs,
    "trHeight",
    "hRule",
    writtenRule === "exact" ? "exact" : "atLeast"
  );
  const child = elementXml(current.tag ?? wName("trHeight"), nextAttrs);
  const rendered = renderProps(setChild(props, "trHeight", child));
  if (rendered === trPr) return null;
  const parsed = parsePropsXml(rendered);
  const format = parsed ? readRowFormat(parsed) : null;
  return format ? { trPr: rendered, format } : null;
}
