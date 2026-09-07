// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import {
  NO_FILL,
  RUN_FORMAT_KEYS,
  type RunFormat,
  toRunFormat,
} from "../../model/format";
import { childOrderOf } from "../../ooxml/childOrder";
import { renderProps } from "../../ooxml/props";
import { editRunProps, type RunProps, readRunProps } from "../runProps";
import {
  type ChildEdit,
  EDITABLE_RUN_KEYS,
  type EditableRunKey,
  EMPTY_RUN_PROPS,
  RUN_PROPERTIES,
  type RunEdit,
  type RunEditContext,
  type RunSetting,
  type RunSettings,
  rPrOf,
  runChildEdits,
  withChildEdits,
} from "./runProperties";

/** One value each editable property can be set to, and the display value it then reads as */
const SAMPLE: {
  [K in EditableRunKey]: { value: RunSetting<K>; reads: RunFormat[K] };
} = {
  bold: { value: true, reads: true },
  italic: { value: true, reads: true },
  strike: { value: true, reads: true },
  doubleStrike: { value: true, reads: true },
  smallCaps: { value: true, reads: true },
  caps: { value: true, reads: true },
  underline: { value: "double", reads: "double" },
  fontSizePt: { value: 10.5, reads: 10.5 },
  characterSpacingPt: { value: -1.5, reads: -1.5 },
  fontFamily: { value: "Arial", reads: '"Arial"' },
  color: { value: "#2e74b5", reads: "#2E74B5" },
  highlight: { value: "yellow", reads: "yellow" },
  background: { value: "#fff2cc", reads: "#FFF2CC" },
  verticalAlign: { value: "superscript", reads: "superscript" },
};

/** Everything a layer below the run could lay down, so that every withdrawal has something to pin against */
const EVERYTHING_BELOW: RunFormat = {
  bold: true,
  italic: true,
  strike: true,
  doubleStrike: true,
  smallCaps: true,
  caps: true,
  underline: "single",
  fontSizePt: 12,
  characterSpacingPt: 1,
  fontFamily: '"Batang"',
  color: "#FF0000",
  highlight: "green",
  background: "#00FF00",
  verticalAlign: "subscript",
  lang: "ko-KR",
};

const NOTHING_BELOW: RunEditContext = { rPr: EMPTY_RUN_PROPS, inherited: {} };

function edit<K extends EditableRunKey>(
  key: K,
  value: RunSetting<K> | null
): RunEdit<K> {
  return { key, value };
}

function editsOf<K extends EditableRunKey>(
  key: K,
  value: RunSetting<K> | null,
  inherited: RunFormat = {}
): ChildEdit[] {
  const edits = runChildEdits(edit(key, value), {
    rPr: EMPTY_RUN_PROPS,
    inherited,
  });
  if (edits === null) throw new Error(`${key} refused ${String(value)}`);
  return edits;
}

function matches<K extends EditableRunKey>(
  key: K,
  format: RunFormat | null,
  value: RunSetting<K>
): boolean {
  return RUN_PROPERTIES[key].matches(format ?? {}, value);
}

/** An rPr holding nothing but the one property */
function written<K extends EditableRunKey>(
  key: K,
  value: RunSetting<K>
): string {
  return renderProps(withChildEdits(EMPTY_RUN_PROPS, editsOf(key, value)));
}

describe("the table", () => {
  it("names every key of RunFormat, and no other", () => {
    expect(Object.keys(RUN_PROPERTIES).sort()).toEqual(
      [...RUN_FORMAT_KEYS].sort()
    );
  });

  it.each(RUN_FORMAT_KEYS)(
    "%s owns children standing in the order CT_RPr lays down",
    (key) => {
      const order = childOrderOf("rPr");
      const children = RUN_PROPERTIES[key].children;
      const at = children.map((child) => order.indexOf(child));
      expect(children.length).toBeGreaterThan(0);
      expect(at).not.toContain(-1);
      expect(at).toEqual([...at].sort((a, b) => a - b));
    }
  );

  it("background owns the shading and the highlight together", () => {
    // Word paints the highlight over the shading, so neither can move without the other
    expect(RUN_PROPERTIES.background.children).toEqual(["highlight", "shd"]);
    expect(editsOf("background", "#FF0000", { highlight: "yellow" })).toEqual([
      ["shd", '<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>'],
      ["highlight", '<w:highlight w:val="none"/>'],
    ]);
  });
});

describe("every editable property", () => {
  it.each(EDITABLE_RUN_KEYS)("%s reads back the value it wrote", (key) => {
    const { value, reads } = SAMPLE[key];
    expect(readRunProps(written(key, value))).toEqual({ [key]: reads });
  });

  it.each(EDITABLE_RUN_KEYS)(
    "%s reads its own write as already in that state",
    (key) => {
      const { value } = SAMPLE[key];
      const format = readRunProps(written(key, value));
      expect(matches(key, format, value)).toBe(true);
      expect(RUN_PROPERTIES[key].isOn(format ?? {})).toBe(true);
      expect(matches(key, null, value)).toBe(false);
    }
  );

  it.each(EDITABLE_RUN_KEYS)(
    "%s writes and withdraws only the children it owns",
    (key) => {
      const owned = RUN_PROPERTIES[key].children;
      for (const inherited of [{}, EVERYTHING_BELOW]) {
        for (const edits of [
          editsOf(key, SAMPLE[key].value, inherited),
          editsOf(key, null, inherited),
        ]) {
          for (const [name] of edits) expect(owned).toContain(name);
        }
      }
    }
  );

  it.each(RUN_FORMAT_KEYS)(
    "what %s reads comes back in through toRunFormat unchanged",
    (key) => {
      // A display value leaves for the screen as data-fmt and is read back through the validator
      const value: RunFormat = { [key]: EVERYTHING_BELOW[key] };
      expect(toRunFormat(value)).toEqual(value);
    }
  );
});

describe("withdrawing a property", () => {
  const pin = <K extends EditableRunKey>(
    key: K,
    inherited: RunFormat,
    pinned: ChildEdit[]
  ) => [key, inherited, pinned] as const;

  it.each([
    pin("bold", { bold: true }, [
      ["b", '<w:b w:val="0"/>'],
      ["bCs", '<w:bCs w:val="0"/>'],
    ]),
    pin("italic", { italic: true }, [
      ["i", '<w:i w:val="0"/>'],
      ["iCs", '<w:iCs w:val="0"/>'],
    ]),
    pin("strike", { strike: true }, [["strike", '<w:strike w:val="0"/>']]),
    pin("doubleStrike", { doubleStrike: true }, [
      ["dstrike", '<w:dstrike w:val="0"/>'],
    ]),
    pin("smallCaps", { smallCaps: true }, [
      ["smallCaps", '<w:smallCaps w:val="0"/>'],
    ]),
    pin("caps", { caps: true }, [["caps", '<w:caps w:val="0"/>']]),
    pin("underline", { underline: "single" }, [["u", '<w:u w:val="none"/>']]),
    pin("color", { color: "#FF0000" }, [["color", '<w:color w:val="auto"/>']]),
    pin("highlight", { highlight: "yellow" }, [
      ["highlight", '<w:highlight w:val="none"/>'],
    ]),
    pin("background", { background: "#FFFF00" }, [
      ["shd", '<w:shd w:val="clear" w:color="auto" w:fill="auto"/>'],
      ["highlight", null],
    ]),
    pin("background", { highlight: "yellow" }, [
      ["shd", null],
      ["highlight", '<w:highlight w:val="none"/>'],
    ]),
    pin("verticalAlign", { verticalAlign: "subscript" }, [
      ["vertAlign", '<w:vertAlign w:val="baseline"/>'],
    ]),
  ])(
    "pins %s off where the layers below switch it on",
    (key, inherited, pinned) => {
      expect(editsOf(key, null, inherited)).toEqual(pinned);
    }
  );

  it.each([
    pin("bold", { bold: false }, [
      ["b", null],
      ["bCs", null],
    ]),
    pin("underline", { underline: "none" }, [["u", null]]),
    pin("background", { background: NO_FILL }, [
      ["shd", null],
      ["highlight", null],
    ]),
  ])(
    "takes %s away where the layers below leave it off anyway",
    (key, inherited, removed) => {
      expect(editsOf(key, null, inherited)).toEqual(removed);
    }
  );

  it.each(EDITABLE_RUN_KEYS)(
    "takes %s away where nothing below says anything",
    (key) => {
      for (const [, xml] of editsOf(key, null)) expect(xml).toBeNull();
    }
  );

  it.each(["fontSizePt", "characterSpacingPt", "fontFamily"] as const)(
    "%s has no off, so a withdrawal takes it away whatever stands below",
    (key) => {
      for (const [, xml] of editsOf(key, null, EVERYTHING_BELOW)) {
        expect(xml).toBeNull();
      }
    }
  );
});

describe("reading the properties that arrived with the table", () => {
  it("reads caps, double strikethrough and character spacing", () => {
    expect(
      readRunProps(
        '<w:rPr><w:caps/><w:dstrike w:val="0"/><w:spacing w:val="30"/></w:rPr>'
      )
    ).toEqual({ caps: true, doubleStrike: false, characterSpacingPt: 1.5 });
  });

  it("reads a condensed spacing and one written as a length", () => {
    expect(readRunProps('<w:rPr><w:spacing w:val="-20"/></w:rPr>')).toEqual({
      characterSpacingPt: -1,
    });
    expect(readRunProps('<w:rPr><w:spacing w:val="1.5pt"/></w:rPr>')).toEqual({
      characterSpacingPt: 1.5,
    });
  });

  it("leaves a spacing that is no measurement out", () => {
    expect(readRunProps('<w:rPr><w:spacing w:val="wide"/></w:rPr>')).toEqual(
      {}
    );
  });

  it("writes a spacing to the nearest twip, the way paragraph spacing is written", () => {
    expect(editsOf("characterSpacingPt", 0.333)).toEqual([
      ["spacing", '<w:spacing w:val="7"/>'],
    ]);
  });

  it("refuses to write a spacing that is no number", () => {
    expect(
      runChildEdits(edit("characterSpacingPt", Number.NaN), NOTHING_BELOW)
    ).toBeNull();
  });
});

describe("rPrOf", () => {
  const settings: RunSettings = {
    bold: true,
    underline: "single",
    fontSizePt: 11,
    fontFamily: "Malgun Gothic",
    color: "#2e74b5",
    background: "#fff2cc",
  };

  it("writes one rPr equal to applying the edits one at a time", () => {
    const edits: RunEdit[] = [
      { key: "bold", value: true },
      { key: "underline", value: "single" },
      { key: "fontSizePt", value: 11 },
      { key: "fontFamily", value: "Malgun Gothic" },
      { key: "color", value: "#2e74b5" },
      { key: "background", value: "#fff2cc" },
    ];
    const oneAtATime = edits.reduce<RunProps>(
      (props, next) => editRunProps(props, {}, next) ?? props,
      { rPr: null, format: null }
    );
    expect(rPrOf(settings)).toBe(oneAtATime.rPr);
  });

  it("writes every setting into one rPr in the order CT_RPr lays down", () => {
    expect(rPrOf(settings)).toBe(
      '<w:rPr><w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="Malgun Gothic"/>' +
        '<w:b/><w:bCs/><w:color w:val="2E74B5"/><w:sz w:val="22"/><w:szCs w:val="22"/>' +
        '<w:u w:val="single"/><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:rPr>'
    );
  });

  it("is nothing for settings that record nothing", () => {
    expect(rPrOf({})).toBeNull();
  });

  it("leaves out a setting the document cannot record and keeps the rest", () => {
    expect(rPrOf({ bold: true, fontFamily: 'Ari"al' })).toBe(
      "<w:rPr><w:b/><w:bCs/></w:rPr>"
    );
  });
});
