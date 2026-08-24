import { type ParagraphFormat, toParagraphFormat } from "../../model/format";
import {
  layerTabStopDirectives,
  type TabStopDirective,
} from "../../model/tabStops";

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

const ALIGNMENTS = new Set([
  "start",
  "center",
  "end",
  "decimal",
  "num",
  "bar",
  "clear",
]);

const LEADERS = new Set([
  "none",
  "dot",
  "hyphen",
  "underscore",
  "heavy",
  "middleDot",
]);

function toTabStopDirective(value: unknown): TabStopDirective | null {
  if (
    !isRecord(value) ||
    typeof value.positionPt !== "number" ||
    !Number.isFinite(value.positionPt) ||
    typeof value.align !== "string" ||
    !ALIGNMENTS.has(value.align)
  ) {
    return null;
  }
  if (value.align === "clear") {
    return { positionPt: value.positionPt, align: "clear" };
  }
  const align = value.align as Exclude<TabStopDirective["align"], "clear">;
  const leader =
    typeof value.leader === "string" && LEADERS.has(value.leader)
      ? (value.leader as Exclude<
          TabStopDirective,
          { align: "clear" }
        >["leader"])
      : undefined;
  return {
    positionPt: value.positionPt,
    align,
    ...(leader === undefined ? {} : { leader }),
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
