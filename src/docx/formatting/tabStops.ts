import { type ParagraphFormat, toParagraphFormat } from "../../model/format";
import {
  layerTabStopDirectives,
  type TabStopDirective,
} from "../../model/tabStops";
import { ST_TabJc, ST_TabTlc } from "../../ooxml/simpleTypes";

export type ParagraphFormatLayer = Omit<
  ParagraphFormat,
  "numbering" | "tabStops"
> & {
  numbering?: ParagraphFormat["numbering"] | null;
  tabStops?: readonly TabStopDirective[];
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function toTabStopDirective(value: unknown): TabStopDirective | null {
  if (
    !isRecord(value) ||
    typeof value.positionPt !== "number" ||
    !Number.isFinite(value.positionPt)
  ) {
    return null;
  }
  const align =
    typeof value.align === "string" ? ST_TabJc.parse(value.align) : null;
  if (align === null) return null;
  if (align === "clear") {
    return { positionPt: value.positionPt, align: "clear" };
  }
  const leader =
    typeof value.leader === "string" ? ST_TabTlc.parse(value.leader) : null;
  return {
    positionPt: value.positionPt,
    align,
    ...(leader === null ? {} : { leader }),
  };
}

export function toParagraphFormatLayer(
  value: unknown
): ParagraphFormatLayer | null {
  if (!isRecord(value)) return null;
  const format: ParagraphFormatLayer = toParagraphFormat(value) ?? {};
  if (!Array.isArray(value.tabStops)) return format;
  const tabStops = value.tabStops.flatMap((entry) => {
    const stop = toTabStopDirective(entry);
    return stop === null ? [] : [stop];
  });
  if (tabStops.length > 0) format.tabStops = tabStops;
  return format;
}

export function layerParagraphValues(
  base: ParagraphFormatLayer,
  over: ParagraphFormatLayer
): ParagraphFormatLayer {
  const result: ParagraphFormatLayer = { ...base, ...over };
  const tabStops = layerTabStopDirectives(
    base.tabStops ?? [],
    over.tabStops ?? []
  );
  if (tabStops.length > 0) result.tabStops = tabStops;
  else delete result.tabStops;
  return result;
}
