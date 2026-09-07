/**
 * The simple types the schema names, one `{parse, format}` pair each.
 *
 * A measurement in OOXML is a union rather than a count. `w:pgSz w:w` is either a number of twips
 * or a universal measure such as `8.5in` (§22.9.2.14, §22.9.2.15), and reading only the leading
 * digits turns Letter paper into a page 8.5 twips wide. A boolean is any of `1`, `0`, `true`,
 * `false`, `on` and `off` (§17.17.4). Deciding either at the call site is how a document comes to
 * be read one way in one reader and another way in the next, so both are decided here, and the
 * writer that puts a value back reaches for the same pair.
 *
 * Readers accept the decimal counts some producers write in integer slots. `format` answers null
 * for a value it cannot record; a caller that intends to round must do so before formatting.
 */

import {
  TAB_LEADERS,
  TAB_STOP_ALIGNMENTS,
  type TabAlignment,
  type TabLeader,
} from "../model/tabStops";

export interface SimpleType<T> {
  /** null for an absent or unreadable value; see the type's producer compatibility rules */
  parse(value: string | null): T | null;
  /** null for a value the type cannot record */
  format(value: T): string | null;
}

/** A twip is a twentieth of a point */
export const TWIPS_PER_PT = 20;

export const TWIPS_PER_INCH = 1440;

/** `w:sz` counts a font size in half-points, so 24 is 12pt */
export const HALF_POINTS_PER_PT = 2;

/** `w:sz` on a border counts its thickness in eighths of a point */
export const EIGHTHS_PER_PT = 8;

/**
 * The largest font size a writer here records.
 *
 * The schema puts no ceiling on `w:sz`, but Word's own limit is 819pt, and a size beyond it is
 * one no word processor will show back.
 */
export const MAX_FONT_SIZE_PT = 819;

/** The twips one unit of each suffix ST_UniversalMeasure admits stands for. A pica is 12 points */
const TWIPS_PER_UNIT: Readonly<Record<string, number>> = {
  mm: TWIPS_PER_INCH / 25.4,
  cm: TWIPS_PER_INCH / 2.54,
  in: TWIPS_PER_INCH,
  pt: TWIPS_PER_PT,
  pc: TWIPS_PER_PT * 12,
  pi: TWIPS_PER_PT * 12,
};

/** §22.9.2.15 `-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)` */
const UNIVERSAL_MEASURE = /^(-?[0-9]+(?:\.[0-9]+)?)(mm|cm|in|pt|pc|pi)$/;

/** `8.5in` -> 12240. null for anything ST_UniversalMeasure does not admit (§22.9.2.15) */
export function universalMeasureToTwips(value: string): number | null {
  const matched = UNIVERSAL_MEASURE.exec(value.trim());
  if (!matched) return null;
  const amount = Number(matched[1]);
  const twips = amount * TWIPS_PER_UNIT[matched[2]];
  return Number.isFinite(twips) ? twips : null;
}

/**
 * A number as the counting branch of a measure writes it.
 *
 * The schema counts in integers, but a producer does not always: Google Docs writes
 * `w:tblW w:w="9026.0"` and cell margins with the same trailing zero. Refusing those would drop a
 * width the file plainly states. Measurement readers retain fractions; count readers truncate
 * them. Writers must choose how to quantize a measurement before formatting its integer count.
 * Signed integer types allow an optional plus sign; unsigned counts retain the previous reader's
 * tolerance for that spelling even though XML Schema 1.0 unsignedLong excludes it.
 */
const DECIMAL_NUMBER = /^[+-]?[0-9]+(?:\.[0-9]+)?$/;

function decimalNumber(value: string): number | null {
  const text = value.trim();
  if (!DECIMAL_NUMBER.test(text)) return null;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : null;
}

interface MeasureShape {
  /**
   * The twips one unit of this measure stands for, which is what a universal measure is converted
   * through: 1 for a twip, 10 for a half-point, 2.5 for an eighth of a point.
   */
  twipsPerUnit: number;
  /** Whether the type admits a value below zero */
  signed: boolean;
}

/**
 * A measurement type: a count in its own unit, or a universal measure converted into that unit.
 *
 * `format` writes the count back. A value carrying a fraction is refused rather than rounded,
 * because a measurement one step off the unit the document counts in is a caller's mistake to see
 * rather than one to bury.
 */
function measure({ twipsPerUnit, signed }: MeasureShape): SimpleType<number> {
  const admits = (value: number) => signed || value >= 0;
  return {
    parse(value) {
      if (value === null) return null;
      const counted = decimalNumber(value);
      if (counted !== null) return admits(counted) ? counted : null;
      const twips = universalMeasureToTwips(value);
      if (twips === null) return null;
      const converted = twips / twipsPerUnit;
      return admits(converted) ? converted : null;
    },
    format(value) {
      if (!Number.isSafeInteger(value) || !admits(value)) return null;
      return `${value}`;
    },
  };
}

/** Twips, unsigned (§22.9.2.14): `w:pgSz`, `w:trHeight`, `w:spacing/@before`, `w:tcMar` */
export const ST_TwipsMeasure: SimpleType<number> = measure({
  twipsPerUnit: 1,
  signed: false,
});

/** Twips, signed (§17.18.81): `w:ind`, `w:tab/@pos`, `w:spacing/@line` */
export const ST_SignedTwipsMeasure: SimpleType<number> = measure({
  twipsPerUnit: 1,
  signed: true,
});

const HPS_MEASURE = measure({
  twipsPerUnit: TWIPS_PER_PT / HALF_POINTS_PER_PT,
  signed: false,
});

/**
 * Half-points (§17.18.42): `w:sz` and `w:szCs`.
 *
 * `format` refuses a size no writer here records: one off the half-point step the unit counts in,
 * one of nothing at all, and one past `MAX_FONT_SIZE_PT`.
 */
export const ST_HpsMeasure: SimpleType<number> = {
  parse: HPS_MEASURE.parse,
  format(value) {
    if (value <= 0 || value > MAX_FONT_SIZE_PT * HALF_POINTS_PER_PT) {
      return null;
    }
    return HPS_MEASURE.format(value);
  },
};

/** Half-points, signed (§17.18.80): `w:position` */
export const ST_SignedHpsMeasure: SimpleType<number> = measure({
  twipsPerUnit: TWIPS_PER_PT / HALF_POINTS_PER_PT,
  signed: true,
});

function wholeNumber(signed: boolean): SimpleType<number> {
  const admits = (value: number) => signed || value >= 0;
  return {
    parse(value) {
      if (value === null) return null;
      const parsed = decimalNumber(value);
      if (parsed === null) return null;
      // A count is written whole; one that arrived with a fraction is read as the whole it names
      const whole = Math.trunc(parsed);
      return Number.isSafeInteger(whole) && admits(whole) ? whole : null;
    },
    format(value) {
      if (!Number.isSafeInteger(value) || !admits(value)) return null;
      return `${value}`;
    },
  };
}

/**
 * Eighths of a point (§17.18.23): the thickness of a border side.
 *
 * Alone among the measurements this restricts `ST_UnsignedDecimalNumber` rather than uniting with a
 * universal measure, so a thickness is a count and a unit written beside it is not one.
 */
export const ST_EighthPointMeasure: SimpleType<number> = wholeNumber(false);

/** A whole number, signed (§17.18.10): `w:numId`, `w:ilvl`, `w:start` */
export const ST_DecimalNumber: SimpleType<number> = wholeNumber(true);

/** A whole number, unsigned (§22.9.2.16): `w:id` on a comment or a note */
export const ST_UnsignedDecimalNumber: SimpleType<number> = wholeNumber(false);

/**
 * A boolean (§17.17.4).
 *
 * The schema unites `xsd:boolean` with the pair `on` and `off`, so all six spellings say the same
 * thing and nothing else says either. An element that carries no `w:val` at all is on, which is a
 * rule about the element rather than about this value, so it is the caller that applies it.
 */
export const ST_OnOff: SimpleType<boolean> = {
  parse(value) {
    if (value === null) return null;
    const text = value.trim();
    if (text === "1" || text === "true" || text === "on") return true;
    if (text === "0" || text === "false" || text === "off") return false;
    return null;
  },
  format(value) {
    return value ? "1" : "0";
  },
};

/** A color, which is either six hex digits or the word `auto` (§17.18.38) */
export type HexColor = { kind: "auto" } | { kind: "rgb"; hex: string };

const HEX_RGB = /^[0-9a-fA-F]{6}$/;

/**
 * `ST_HexColor` keeps the digits exactly as the document spelled them, so a color read and written
 * back leaves the file as it was. A caller comparing two colors gathers them first.
 */
export const ST_HexColor: SimpleType<HexColor> = {
  parse(value) {
    if (value === null) return null;
    if (value === "auto") return { kind: "auto" };
    return HEX_RGB.test(value) ? { kind: "rgb", hex: value } : null;
  },
  format(value) {
    return value.kind === "auto" ? "auto" : value.hex;
  },
};

/**
 * The width of a table or a cell (§17.18.107), which `w:type` beside it says how to read.
 *
 * `number` is the bare count, meaning twips under `dxa` and fiftieths of a percent under `pct`.
 * `percent` carried its own `%`, and `twips` was written as a universal measure, which says an
 * absolute width whatever `w:type` claims.
 */
export type MeasurementOrPercent =
  | { kind: "number"; value: number }
  | { kind: "percent"; value: number }
  | { kind: "twips"; value: number };

const PERCENT = /^(-?[0-9]+(?:\.[0-9]+)?)%$/;

export const ST_MeasurementOrPercent: SimpleType<MeasurementOrPercent> = {
  parse(value) {
    if (value === null) return null;
    const text = value.trim();
    const percent = PERCENT.exec(text);
    if (percent) {
      const amount = decimalNumber(percent[1]);
      return amount === null ? null : { kind: "percent", value: amount };
    }
    const counted = decimalNumber(text);
    if (counted !== null) return { kind: "number", value: counted };
    const twips = universalMeasureToTwips(text);
    return twips === null ? null : { kind: "twips", value: twips };
  },
  format(value) {
    if (!Number.isFinite(value.value)) return null;
    if (value.kind === "percent") return `${value.value}%`;
    if (value.kind === "twips") return `${value.value / TWIPS_PER_PT}pt`;
    return ST_DecimalNumber.format(value.value);
  },
};

/** The kinds of custom tab stop, with the two Transitional spellings folded in */
export type TabJc = TabAlignment | "bar" | "clear";

/**
 * The kind of a custom tab stop (§17.18.84).
 *
 * `left` and `right` are Transitional spellings of `start` and `end` (Part 4 §14.11.6), so they
 * are read as those and never written back.
 */
const TAB_JC_BY_VALUE: Readonly<Record<string, TabJc>> = {
  ...Object.fromEntries(TAB_STOP_ALIGNMENTS.map((align) => [align, align])),
  clear: "clear",
  left: "start",
  right: "end",
};

export const ST_TabJc: SimpleType<TabJc> = {
  parse(value) {
    return value !== null && Object.hasOwn(TAB_JC_BY_VALUE, value)
      ? TAB_JC_BY_VALUE[value]
      : null;
  },
  format(value) {
    return value;
  },
};

/** The character a tab stop draws the space it covers with (§17.18.85) */
export const ST_TabTlc: SimpleType<TabLeader> = {
  parse(value) {
    return TAB_LEADERS.find((leader) => leader === value) ?? null;
  },
  format(value) {
    return value;
  },
};
