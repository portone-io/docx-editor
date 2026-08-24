import { type ParagraphFormat, toParagraphFormat } from "../../model/format";
import { mergeTabStops } from "../../model/tabStops";
import {
  EMPTY_NUMBERING,
  levelIndentPt,
  type Numbering,
} from "../../numbering/parseNumbering";
import { parsePropsXml } from "../propsXml";
import { readParagraphFormat } from "./direct";
import {
  paragraphStyleFormat,
  type StyleFormat,
  type StyleTable,
} from "./styles";
import { layerParagraphValues, type ParagraphFormatLayer } from "./tabStops";

export interface ParagraphFormattingContext {
  styles: StyleTable;
  defaultStyleId: string | null;
  defaults: ParagraphFormatLayer;
  numbering: Numbering;
}

export const NO_PARAGRAPH_FORMATTING: ParagraphFormattingContext = {
  styles: new Map(),
  defaultStyleId: null,
  defaults: {},
  numbering: EMPTY_NUMBERING,
};

function numberingRef(
  defaults: ParagraphFormatLayer,
  style: ParagraphFormatLayer,
  direct: ParagraphFormatLayer
): ParagraphFormat["numbering"] | null | undefined {
  let value = defaults.numbering;
  if (style.numbering !== undefined) value = style.numbering;
  if (direct.numbering !== undefined) value = direct.numbering;
  return value;
}

function numberingLayer(
  ref: ParagraphFormat["numbering"] | null | undefined,
  numbering: Numbering
): ParagraphFormatLayer {
  if (!ref) return {};
  const level = numbering.lists.get(ref.numId)?.levels.get(ref.ilvl);
  if (!level) return {};
  const indent = levelIndentPt(level.indent);
  const leadingIndentPt = indent.startPt;
  const implicitHangingStop =
    leadingIndentPt !== null &&
    indent.textIndentPt !== null &&
    indent.textIndentPt < 0
      ? [{ positionPt: leadingIndentPt, align: "start" as const }]
      : [];
  return {
    tabStops: [...implicitHangingStop, ...(level.tabStops ?? [])],
  };
}

function withImplicitHangingStop(
  format: ParagraphFormat | null
): ParagraphFormat | null {
  const leadingIndentPt = format?.indentStartPt ?? format?.indentLeftPt;
  if (
    format?.textIndentPt === undefined ||
    format.textIndentPt >= 0 ||
    leadingIndentPt === undefined
  ) {
    return format;
  }
  const implicit = {
    positionPt: leadingIndentPt,
    align: "start" as const,
  };
  return {
    ...format,
    tabStops: mergeTabStops([implicit], format.tabStops ?? []),
  };
}

function directLayer(pPr: unknown): ParagraphFormatLayer {
  if (typeof pPr !== "string") return {};
  return readParagraphFormat(parsePropsXml(pPr)) ?? {};
}

function styleFor(
  pPr: unknown,
  context: ParagraphFormattingContext
): StyleFormat | undefined {
  return paragraphStyleFormat(pPr, context.styles, context.defaultStyleId);
}

/** Resolves the effective OOXML paragraph-property hierarchy for display. */
export function effectiveParagraphFormat(
  pPr: unknown,
  context: ParagraphFormattingContext
): ParagraphFormat | null {
  const direct = directLayer(pPr);
  const style = styleFor(pPr, context)?.paragraph ?? {};
  const level = numberingLayer(
    numberingRef(context.defaults, style, direct),
    context.numbering
  );
  const layered = [level, style, direct].reduce(
    layerParagraphValues,
    context.defaults
  );
  const format = toParagraphFormat(layered);
  return withImplicitHangingStop(
    format !== null && Object.keys(format).length === 0 ? null : format
  );
}

/** Resolves the run properties supplied by the paragraph style. */
export function effectiveParagraphStyle(
  pPr: unknown,
  context: ParagraphFormattingContext
): StyleFormat | undefined {
  return styleFor(pPr, context);
}
