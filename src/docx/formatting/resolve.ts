/**
 * Resolves the §17.7.2 hierarchy for one paragraph and for the runs inside it.
 *
 * The layers are put in order here and nowhere else, so a paragraph's display values are the
 * same whichever path asked for them: opening the document, an edit, a paste, or a paragraph an
 * edit built from nothing.
 */

import {
  type ParagraphFormat,
  type RunFormat,
  toParagraphFormat,
} from "../../model/format";
import type { TabStop, TabStopDirective } from "../../model/tabStops";
import { listFor } from "../../numbering/listTemplate";
import {
  type LevelIndentPt,
  levelIndentPt,
  type NumberingLevel,
} from "../../numbering/parseNumbering";
import { parsePropsXml } from "../../ooxml/props";
import type {
  ConditionalTableFormat,
  TableStyleOverrideType,
} from "../tableFormatting";
import type { FormattingContext } from "./context";
import { readParagraphFormat, readRunFormat, runStyleIdOf } from "./direct";
import {
  layerRunFormat,
  paragraphStyleFormat,
  type StyleFormat,
} from "./styles";
import type { ParagraphFormatLayer } from "./tabStops";

/** Where a paragraph sits when it is inside a table cell. null for body text */
export interface ParagraphPlacement {
  /** What the table's `w:tblStyle` names. null for a table naming none, which wears the default one */
  tableStyleId: string | null;
  /** The parts of the table the cell belongs to, lowest first (`docx/tableFormatting/conditions`) */
  conditions: readonly TableStyleOverrideType[];
}

/**
 * The layer of the hierarchy a tab stop came from. Removing an inherited stop means writing a
 * `clear` over it, while a direct one is deleted from the paragraph's own list, which is what a
 * ruler needs to tell apart. The document defaults and a table style count as `style`, since a
 * stop from either is taken away the same way
 */
export type TabStopLayer = "numbering" | "style" | "direct" | "implicit";

export interface LayeredTabStop extends TabStop {
  layer: TabStopLayer;
}

export interface ResolvedParagraph {
  /** `format` in `schema`: docDefaults < table style < numbering level < paragraph style < direct */
  format: ParagraphFormat | null;
  /** `styleRun` in `schema`: run defaults < table style rPr < paragraph style rPr. Default font and size stay in CSS variables */
  styleRun: RunFormat | null;
  /** Everything below a run's own formatting, docDefaults included. What a removed run property falls back to */
  inheritedRun: RunFormat;
  /** The stops `format.tabStops` holds, each with the layer that laid it down */
  tabStops: readonly LayeredTabStop[];
}

interface Layer {
  values: ParagraphFormatLayer;
  stops: TabStopLayer;
}

function directLayer(pPr: string | null): ParagraphFormatLayer {
  return pPr === null ? {} : (readParagraphFormat(parsePropsXml(pPr)) ?? {});
}

/**
 * The table style an object wearing this one is dressed by.
 *
 * A table that points at no style, or at one that is not defined, falls back on the document's
 * default table style, which is what OOXML applies to an object with no style of its own.
 */
export function tableStyleFor(
  styleId: string | null,
  context: FormattingContext
): StyleFormat | undefined {
  const named = styleId === null ? undefined : context.styles.get(styleId);
  const fallback =
    context.defaultTableStyleId === null
      ? undefined
      : context.styles.get(context.defaultTableStyleId);
  return named ?? fallback;
}

/**
 * What the table style dresses the parts this cell belongs to with, lowest first.
 * A paragraph in a header row wears what the style wrote for the whole table and then what it
 * wrote for the header row, in the order §17.7.6 lays them over one another.
 */
function conditionsOf(
  style: StyleFormat | undefined,
  placement: ParagraphPlacement | null
): ConditionalTableFormat[] {
  if (!style || !placement) return [];
  return placement.conditions.flatMap((type) => {
    const format = style.tableConditions[type];
    return format ? [format] : [];
  });
}

/** The list slot the layers settle on. `numId` 0 at any layer takes away the one below it (§17.9.18) */
function numberingLevelOf(
  layers: readonly ParagraphFormatLayer[],
  context: FormattingContext
): NumberingLevel | undefined {
  const ref = layers.reduce<ParagraphFormat["numbering"] | null | undefined>(
    (found, layer) => (layer.numbering === undefined ? found : layer.numbering),
    undefined
  );
  if (!ref) return undefined;
  return listFor(context.numbering, ref.numId)?.levels.get(ref.ilvl);
}

/**
 * The level's indent is drawn as a decoration (`numberingDecorations`) and stays out of
 * `format`, so a list paragraph's own indent can still be told from the one it inherits.
 */
function numberingLayer(
  level: NumberingLevel | undefined
): ParagraphFormatLayer {
  return level?.tabStops ? { tabStops: level.tabStops } : {};
}

function layerValues(
  base: ParagraphFormatLayer,
  over: ParagraphFormatLayer
): ParagraphFormatLayer {
  const { tabStops: _stops, ...values } = over;
  return { ...base, ...values };
}

function byPosition(left: TabStop, right: TabStop): number {
  return left.positionPt - right.positionPt;
}

/**
 * The stops the layers leave standing (§17.3.1.38: additive, a `clear` taking away the stop
 * inherited at its position), each remembering the layer that last spoke for its position.
 */
function explicitStops(layers: readonly Layer[]): LayeredTabStop[] {
  const spoken = new Map<
    number,
    { stop: TabStopDirective; layer: TabStopLayer }
  >();
  for (const { values, stops } of layers) {
    for (const stop of values.tabStops ?? []) {
      spoken.set(stop.positionPt, { stop, layer: stops });
    }
  }
  return [...spoken.values()]
    .flatMap(({ stop, layer }) =>
      stop.align === "clear" ? [] : [{ ...stop, layer }]
    )
    .sort(byPosition);
}

/**
 * Where a hanging indent puts the stop §17.3.1.38 says it always creates. The indent a list
 * level lays down never enters `format`, so it is read here the way the decoration draws it:
 * only where the paragraph wrote none of its own.
 */
function hangingStopPt(
  values: ParagraphFormatLayer,
  level: LevelIndentPt
): number | null {
  const leading = values.indentStartPt ?? values.indentLeftPt ?? level.startPt;
  const textIndent = values.textIndentPt ?? level.textIndentPt;
  return leading !== null && textIndent !== null && textIndent < 0
    ? leading
    : null;
}

/** A stop written at the hanging indent's position takes the implicit one's place */
function withImplicitStop(
  stops: readonly LayeredTabStop[],
  positionPt: number | null
): readonly LayeredTabStop[] {
  if (positionPt === null) return stops;
  if (stops.some((stop) => stop.positionPt === positionPt)) return stops;
  const implicit: LayeredTabStop = {
    positionPt,
    align: "start",
    layer: "implicit",
  };
  return [...stops, implicit].sort(byPosition);
}

function orNull<T extends object>(values: T): T | null {
  return Object.keys(values).length === 0 ? null : values;
}

export function resolveParagraph(
  pPr: string | null,
  context: FormattingContext,
  placement: ParagraphPlacement | null = null
): ResolvedParagraph {
  const tableStyle =
    placement === null
      ? undefined
      : tableStyleFor(placement.tableStyleId, context);
  const conditions = conditionsOf(tableStyle, placement);
  const style = paragraphStyleFormat(
    pPr,
    context.styles,
    context.defaultParagraphStyleId
  );
  const direct = directLayer(pPr);
  const table: Layer[] = [
    { values: tableStyle?.paragraph ?? {}, stops: "style" },
    ...conditions.map(
      (condition): Layer => ({ values: condition.paragraph, stops: "style" })
    ),
  ];
  const explicit: Layer[] = [
    { values: context.paragraphDefaults, stops: "style" },
    ...table,
    { values: style?.paragraph ?? {}, stops: "style" },
    { values: direct, stops: "direct" },
  ];
  const level = numberingLevelOf(
    explicit.map((layer) => layer.values),
    context
  );
  // The numbering level sits above the table style and below the paragraph style (§17.7.2)
  const belowNumbering = 1 + table.length;
  const layers: Layer[] = [
    ...explicit.slice(0, belowNumbering),
    { values: numberingLayer(level), stops: "numbering" },
    ...explicit.slice(belowNumbering),
  ];
  const values = layers.reduce<ParagraphFormatLayer>(
    (base, layer) => layerValues(base, layer.values),
    {}
  );
  const stops = explicitStops(layers);
  const tabStops = context.compat.noTabHangInd
    ? stops
    : withImplicitStop(
        stops,
        hangingStopPt(values, levelIndentPt(level?.indent ?? null))
      );
  const format = toParagraphFormat(values) ?? {};
  if (tabStops.length > 0) {
    format.tabStops = tabStops.map(({ layer: _layer, ...stop }) => stop);
  }
  // Only the default font and size are supplied by the sheet's CSS variables. Other
  // run defaults need display values too, or a bold default is invisible to both text and toolbar.
  const {
    fontSizePt: _size,
    fontFamily: _font,
    ...displayDefaults
  } = context.runDefaults;
  const styleRun: RunFormat = {
    ...conditions.reduce<RunFormat>(
      (run, condition) => ({ ...run, ...condition.run }),
      { ...displayDefaults, ...tableStyle?.run }
    ),
    ...style?.run,
  };
  return {
    format: orNull(format),
    styleRun: orNull(styleRun),
    inheritedRun: { ...context.runDefaults, ...styleRun },
    tabStops,
  };
}

/** The values the character style a run points at lays down (`w:rStyle`). Nothing for a run pointing at none */
function characterStyleRun(
  rPr: Element | null,
  context: FormattingContext
): RunFormat {
  const id = runStyleIdOf(rPr);
  return id === null ? {} : (context.styles.get(id)?.run ?? {});
}

/** `format` on a run mark: paragraph.styleRun < character style < the run's own rPr */
export function resolveRun(
  rPr: string | null,
  paragraph: ResolvedParagraph,
  context: FormattingContext
): RunFormat | null {
  const el = rPr === null ? null : parsePropsXml(rPr);
  return layerRunFormat(
    { ...paragraph.styleRun, ...characterStyleRun(el, context) },
    readRunFormat(el, context.themeFonts)
  );
}

/** paragraph.inheritedRun < character style. What a run property taken out of the rPr falls back to */
export function inheritedRunFormat(
  rPr: string | null,
  paragraph: ResolvedParagraph,
  context: FormattingContext
): RunFormat {
  const el = rPr === null ? null : parsePropsXml(rPr);
  return { ...paragraph.inheritedRun, ...characterStyleRun(el, context) };
}
