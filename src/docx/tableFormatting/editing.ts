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
import { attrPairs, elementXml, type XmlAttr } from "../../ooxml/element";
import { wName } from "../../ooxml/names";
import { setAttr } from "../../ooxml/precedence";
import { normalizeHex, wAttr } from "../../ooxml/units";
import { childByLocalName } from "../../ooxml/xml";
import {
  type Props,
  parseProps,
  parsePropsXml,
  propsChild,
  renderProps,
  setPropsChild,
  TC_PR_ORDER,
  TR_PR_ORDER,
} from "../propsXml";
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

/**
 * The order the children are laid out in under `w:tcBorders` (CT_TcBorders).
 * `start` and `end` are the newer spelling of `left` and `right`, so each pair shares one slot.
 */
const TC_BORDERS_ORDER: readonly string[] = [
  "top",
  "start",
  "left",
  "bottom",
  "end",
  "right",
  "insideH",
  "insideV",
  "tl2br",
  "tr2bl",
];

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
function drawsLine(side: Element | null): boolean {
  if (!side) return false;
  const val = wAttr(side, "val");
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
  current: Element | null,
  edit: BordersEdit,
  fallback: string | null = null
): readonly XmlAttr[] | null {
  if (edit.kind === "color") {
    if (!drawsLine(current)) {
      if (current || !fallback || fallback === "none") return null;
      return inheritedBorderAttrs(fallback, edit.hex);
    }
    return setAttr(attrPairs(current), name, "color", edit.hex ?? "auto");
  }
  if (edit.line === "none") {
    // The thickness and color stay behind, so a line switched off comes back as it was
    return current
      ? setAttr(attrPairs(current), name, "val", "none")
      : [[wName("val"), "none"]];
  }
  if (!current) {
    return [
      [wName("val"), "single"],
      [wName("sz"), `${PRESET_BORDER_EIGHTHS}`],
      [wName("space"), "0"],
      [wName("color"), "auto"],
    ];
  }
  // The color is not ours to decide here, so a themed or explicit color survives the preset
  return setAttr(
    setAttr(attrPairs(current), name, "val", "single"),
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
  borders: Element | null,
  side: CellSide
): { write: string; drop: string | null } {
  const [primary, alternate] = SIDE_NAMES[side];
  const usesAlternate =
    alternate !== undefined &&
    borders !== null &&
    childByLocalName(borders, primary) === null &&
    childByLocalName(borders, alternate) !== null;
  if (usesAlternate) return { write: alternate, drop: null };
  return { write: primary, drop: alternate ?? null };
}

/**
 * The `<w:tcBorders>` fragment with the sides the job names rewritten.
 * null if the original fragment's shape could not be made out, in which case the caller leaves the
 * cell untouched. `xml` is null when no side is left to write, which removes the fragment.
 */
function editedBorders(
  current: string | null,
  edit: BordersEdit,
  defaults: CellBorderDefaults = NO_BORDER_DEFAULTS
): { xml: string | null } | null {
  const props =
    current === null
      ? { tag: "w:tcBorders", attrs: null, children: [] }
      : parseProps(current);
  const borders = current === null ? null : parsePropsXml(current);
  if (!props || (current !== null && !borders)) return null;

  const sides = edit.kind === "line" ? edit.sides : ALL_CELL_SIDES;
  const edited = sides.reduce((kept, side) => {
    const { write, drop } = sideNames(borders, side);
    const currentSide = borders ? childByLocalName(borders, write) : null;
    const attrs = sideAttrs(write, currentSide, edit, defaults[side]);
    if (!attrs) return kept;
    // The original tag keeps its prefix, so a document not using `w:` is written back as it was
    const tag = currentSide?.nodeName ?? wName(write);
    const written = setPropsChild(
      kept,
      write,
      elementXml(tag, attrs),
      TC_BORDERS_ORDER
    );
    return drop === null
      ? written
      : setPropsChild(written, drop, null, TC_BORDERS_ORDER);
  }, props);

  const xml = renderProps(edited);
  return { xml: xml === "" ? null : xml };
}

/** Whether the shading paints a pattern rather than a plain fill */
function hasPattern(shd: Element | null): boolean {
  const val = shd ? wAttr(shd, "val") : null;
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
  current: Element | null,
  fill: string | null
): string | null {
  const tag = current?.nodeName ?? wName("shd");
  if (fill === null) {
    if (!hasPattern(current)) return null;
    return elementXml(tag, setAttr(attrPairs(current), "shd", "fill", "auto"));
  }
  if (!current) {
    return elementXml(tag, [
      [wName("val"), "clear"],
      [wName("color"), "auto"],
      [wName("fill"), fill],
    ]);
  }
  const painting = hasPattern(current)
    ? attrPairs(current)
    : setAttr(attrPairs(current), "shd", "val", "clear");
  return elementXml(tag, setAttr(painting, "shd", "fill", fill));
}

/** What one child of the tcPr is to be changed to. A null xml removes that child */
interface ChildChange {
  name: string;
  xml: string | null;
}

function bordersChange(
  props: Props,
  edit: BordersEdit,
  defaults: CellBorderDefaults
): ChildChange | null {
  const current = propsChild(props.children, "tcBorders")?.xml ?? null;
  const edited = editedBorders(current, edit, defaults);
  return edited === null ? null : { name: "tcBorders", xml: edited.xml };
}

const TC_MAR_ORDER: readonly string[] = [
  "top",
  "start",
  "left",
  "bottom",
  "end",
  "right",
];

const MARGIN_SIDE_NAMES: Readonly<Record<CellSide, readonly string[]>> = {
  top: ["top"],
  right: ["end", "right"],
  bottom: ["bottom"],
  left: ["start", "left"],
};

function paddingChange(
  props: Props,
  values: Partial<Record<CellSide, number>>
): ChildChange | null {
  const current = propsChild(props.children, "tcMar")?.xml ?? null;
  const margins = current === null ? null : parsePropsXml(current);
  const parsed =
    current === null
      ? { tag: "w:tcMar", attrs: null, children: [] }
      : parseProps(current);
  if (!parsed || (current !== null && !margins)) return null;
  let edited = parsed;
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
    const existing = margins
      ? (MARGIN_SIDE_NAMES[side]
          .map((name) => childByLocalName(margins, name))
          .find((element) => element !== null) ?? null)
      : null;
    const name = existing?.localName ?? side;
    const attrs = setAttr(
      setAttr(attrPairs(existing), name, "w", `${twips}`),
      name,
      "type",
      "dxa"
    );
    edited = setPropsChild(
      edited,
      name,
      elementXml(existing?.nodeName ?? wName(side), attrs),
      TC_MAR_ORDER
    );
    wrote = true;
  }
  if (!wrote) return null;
  return { name: "tcMar", xml: renderProps(edited) };
}

/** Which child of the tcPr one job changes and how. null for a value that cannot be written down */
function childChange(
  edit: CellFormatEdit,
  props: Props,
  defaults: CellBorderDefaults
): ChildChange | null {
  switch (edit.kind) {
    case "background": {
      const fill = edit.hex === null ? null : normalizeHex(edit.hex);
      if (edit.hex !== null && fill === null) return null;
      const current = propsChild(props.children, "shd")?.xml ?? null;
      const shd = current === null ? null : parsePropsXml(current);
      if (current !== null && !shd) return null;
      return { name: "shd", xml: editedShading(shd, fill) };
    }
    case "borders":
      return bordersChange(
        props,
        {
          kind: "line",
          line: edit.line,
          sides: edit.sides,
        },
        defaults
      );
    case "borderColor": {
      const hex = edit.hex === null ? null : normalizeHex(edit.hex);
      if (edit.hex !== null && hex === null) return null;
      return bordersChange(props, { kind: "color", hex }, defaults);
    }
    case "verticalAlign": {
      const current = propsChild(props.children, "vAlign")?.xml ?? null;
      const element = current === null ? null : parsePropsXml(current);
      if (current !== null && !element) return null;
      return {
        name: "vAlign",
        xml: elementXml(
          element?.nodeName ?? wName("vAlign"),
          setAttr(attrPairs(element), "vAlign", "val", edit.align)
        ),
      };
    }
    case "padding":
      return paddingChange(props, edit.values);
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

  const change = childChange(edit, props, defaults);
  if (!change) return null;

  const rendered = renderProps(
    setPropsChild(props, change.name, change.xml, TC_PR_ORDER)
  );
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
  const el = tcPr === null ? null : parsePropsXml(tcPr);
  const borders = el ? childByLocalName(el, "tcBorders") : null;
  if (!borders) return false;
  return ALL_CELL_SIDES.some((side) =>
    SIDE_NAMES[side].some((name) => drawsLine(childByLocalName(borders, name)))
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
      ? { tag: "w:trPr", attrs: null, children: [] }
      : parseProps(trPr);
  if (!props) return null;
  const current = propsChild(props.children, "trHeight")?.xml ?? null;
  const element = current === null ? null : parsePropsXml(current);
  if (current !== null && !element) return null;
  const writtenRule = element ? wAttr(element, "hRule") : null;
  const attrs = setAttr(attrPairs(element), "trHeight", "val", `${twips}`);
  const nextAttrs = setAttr(
    attrs,
    "trHeight",
    "hRule",
    writtenRule === "exact" ? "exact" : "atLeast"
  );
  const child = elementXml(element?.nodeName ?? wName("trHeight"), nextAttrs);
  const rendered = renderProps(
    setPropsChild(props, "trHeight", child, TR_PR_ORDER)
  );
  if (rendered === trPr) return null;
  const parsed = parsePropsXml(rendered);
  const format = parsed ? readRowFormat(parsed) : null;
  return format ? { trPr: rendered, format } : null;
}
