// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { readDefaultTabStop } from "../docx/documentSettings";
import {
  type HexColor,
  type MeasurementOrPercent,
  type SimpleType,
  ST_DecimalNumber,
  ST_EighthPointMeasure,
  ST_HexColor,
  ST_HpsMeasure,
  ST_MeasurementOrPercent,
  ST_OnOff,
  ST_SignedHpsMeasure,
  ST_SignedTwipsMeasure,
  ST_TabJc,
  ST_TabTlc,
  ST_TwipsMeasure,
  ST_UnsignedDecimalNumber,
  universalMeasureToTwips,
} from "./simpleTypes";
import { readTabStopDirectives } from "./tabStops";
import { parseXml } from "./xml";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** One type and a value it admits, named so a failing row says which pair broke */
const ROUND_TRIPS: readonly [string, SimpleType<unknown>, unknown][] = [
  ["ST_TwipsMeasure", ST_TwipsMeasure, 12240],
  ["ST_TwipsMeasure zero", ST_TwipsMeasure, 0],
  ["ST_SignedTwipsMeasure", ST_SignedTwipsMeasure, -720],
  ["ST_HpsMeasure", ST_HpsMeasure, 24],
  ["ST_SignedHpsMeasure", ST_SignedHpsMeasure, -12],
  ["ST_EighthPointMeasure", ST_EighthPointMeasure, 4],
  ["ST_DecimalNumber", ST_DecimalNumber, -3],
  ["ST_UnsignedDecimalNumber", ST_UnsignedDecimalNumber, 7],
  ["ST_OnOff on", ST_OnOff, true],
  ["ST_OnOff off", ST_OnOff, false],
  ["ST_HexColor auto", ST_HexColor, { kind: "auto" } satisfies HexColor],
  [
    "ST_HexColor rgb",
    ST_HexColor,
    { kind: "rgb", hex: "2E74B5" } satisfies HexColor,
  ],
  [
    "ST_MeasurementOrPercent number",
    ST_MeasurementOrPercent,
    { kind: "number", value: 9026 } satisfies MeasurementOrPercent,
  ],
  [
    "ST_MeasurementOrPercent percent",
    ST_MeasurementOrPercent,
    { kind: "percent", value: 50 } satisfies MeasurementOrPercent,
  ],
  [
    "ST_MeasurementOrPercent twips",
    ST_MeasurementOrPercent,
    { kind: "twips", value: 12240 } satisfies MeasurementOrPercent,
  ],
  ["ST_TabJc", ST_TabJc, "end"],
  ["ST_TabTlc", ST_TabTlc, "middleDot"],
];

describe("the simple types the schema names", () => {
  it.each(ROUND_TRIPS)("parse(format(v)) === v for %s", (_, type, value) => {
    const written = type.format(value);
    expect(written).not.toBeNull();
    expect(type.parse(written)).toEqual(value);
  });

  it("reads 8.5in as 12240 twips and 12pt as 24 half-points", () => {
    expect(ST_TwipsMeasure.parse("8.5in")).toBe(12240);
    expect(ST_TwipsMeasure.parse("11in")).toBe(15840);
    expect(ST_TwipsMeasure.parse("2.54cm")).toBe(1440);
    expect(ST_TwipsMeasure.parse("25.4mm")).toBe(1440);
    expect(ST_TwipsMeasure.parse("6pc")).toBe(1440);
    expect(ST_TwipsMeasure.parse("6pi")).toBe(1440);
    expect(ST_HpsMeasure.parse("12pt")).toBe(24);
    // A border thickness is the one measurement that unites with no universal measure
    expect(ST_EighthPointMeasure.parse("4")).toBe(4);
    expect(ST_EighthPointMeasure.parse("0.5pt")).toBeNull();
    expect(ST_SignedTwipsMeasure.parse("-0.5in")).toBe(-720);
  });

  it("refuses a measure of a unit no universal measure names", () => {
    expect(ST_TwipsMeasure.parse("8.5em")).toBeNull();
    expect(ST_TwipsMeasure.parse("8.5")).toBe(8.5);
    expect(ST_TwipsMeasure.parse("in")).toBeNull();
    expect(ST_TwipsMeasure.parse("")).toBeNull();
    expect(ST_TwipsMeasure.parse(null)).toBeNull();
    // An unsigned measure has no reading for a value below zero
    expect(ST_TwipsMeasure.parse("-720")).toBeNull();
    expect(ST_TwipsMeasure.parse("-0.5in")).toBeNull();
    expect(universalMeasureToTwips("8.5in")).toBe(12240);
    expect(universalMeasureToTwips("8.5")).toBeNull();
  });

  it("reads the decimal count a producer writes in place of a whole one", () => {
    // Google Docs writes `w:tblW w:w="9026.0"`, and refusing it drops the width the file states
    expect(ST_MeasurementOrPercent.parse("9026.0")).toEqual({
      kind: "number",
      value: 9026,
    });
    expect(ST_TwipsMeasure.parse("108.0")).toBe(108);
    expect(ST_DecimalNumber.parse("2.0")).toBe(2);
  });

  it("reads on and off spellings as the schema admits and rejects the rest", () => {
    for (const on of ["1", "true", "on"]) {
      expect(ST_OnOff.parse(on)).toBe(true);
    }
    for (const off of ["0", "false", "off"]) {
      expect(ST_OnOff.parse(off)).toBe(false);
    }
    for (const neither of ["", "yes", "no", "t", "f", "On", "OFF", "2"]) {
      expect(ST_OnOff.parse(neither)).toBeNull();
    }
    expect(ST_OnOff.parse(null)).toBeNull();
  });

  it("formats a percent width back with the % it came in with", () => {
    const written = ST_MeasurementOrPercent.parse("50%");
    expect(written).toEqual({ kind: "percent", value: 50 });
    expect(written && ST_MeasurementOrPercent.format(written)).toBe("50%");
    // The same width written as fiftieths carries no % and gets none back
    const fiftieths = ST_MeasurementOrPercent.parse("2500");
    expect(fiftieths).toEqual({ kind: "number", value: 2500 });
    expect(fiftieths && ST_MeasurementOrPercent.format(fiftieths)).toBe("2500");
  });

  it("refuses to record a size Word does not keep", () => {
    // A font size counts in half-points, so one off that step is not a size a document can hold
    expect(ST_HpsMeasure.format(24)).toBe("24");
    expect(ST_HpsMeasure.format(24.5)).toBeNull();
    expect(ST_HpsMeasure.format(0)).toBeNull();
    expect(ST_HpsMeasure.format(1638)).toBe("1638");
    expect(ST_HpsMeasure.format(1639)).toBeNull();
    expect(ST_TwipsMeasure.format(-1)).toBeNull();
    expect(ST_TwipsMeasure.format(Number.NaN)).toBeNull();
    expect(ST_TwipsMeasure.format(720.5)).toBeNull();
  });

  it("folds the transitional tab spellings and keeps the leader names", () => {
    expect(ST_TabJc.parse("left")).toBe("start");
    expect(ST_TabJc.parse("right")).toBe("end");
    expect(ST_TabJc.parse("bar")).toBe("bar");
    expect(ST_TabJc.parse("middle")).toBeNull();
    expect(ST_TabTlc.parse("dot")).toBe("dot");
    expect(ST_TabTlc.parse("dots")).toBeNull();
  });

  it("reads a tab position and a default tab interval written as universal measures", () => {
    const pPr = parseXml(
      `<w:pPr ${W_NS}><w:tabs><w:tab w:val="left" w:pos="1.5in"/></w:tabs></w:pPr>`
    ).documentElement;
    expect(readTabStopDirectives(pPr)).toEqual([
      { positionPt: 108, align: "start" },
    ]);

    const settings = parseXml(
      `<w:settings ${W_NS}><w:defaultTabStop w:val="0.75in"/></w:settings>`
    );
    expect(readDefaultTabStop(settings)).toBe(54);
  });
});
