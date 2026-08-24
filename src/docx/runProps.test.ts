// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import {
  editRunProps,
  isRunToggleOn,
  matchesRunEdit,
  type RunEdit,
  type RunProps,
  readRunProps,
} from "./runProps";

/** Changes one piece of formatting inside a paragraph that points at no style */
function edit(rPr: string | null, change: RunEdit): RunProps | null {
  return editRunProps({ rPr, format: readRunProps(rPr) }, null, change);
}

/** Changes one piece of formatting inside a paragraph that points at a style */
function editStyled(rPr: string | null, change: RunEdit): RunProps | null {
  return editRunProps(
    { rPr, format: readRunProps(rPr) },
    '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>',
    change
  );
}

const bold = (on: boolean): RunEdit => ({ kind: "toggle", toggle: "bold", on });
const italic = (on: boolean): RunEdit => ({
  kind: "toggle",
  toggle: "italic",
  on,
});
const underline = (on: boolean): RunEdit => ({
  kind: "toggle",
  toggle: "underline",
  on,
});
const strike = (on: boolean): RunEdit => ({
  kind: "toggle",
  toggle: "strike",
  on,
});

describe("turning formatting on", () => {
  it("creates a minimal rPr for a run that has no formatting", () => {
    expect(edit(null, bold(true))).toEqual({
      rPr: "<w:rPr><w:b/><w:bCs/></w:rPr>",
      format: { bold: true },
    });
  });

  it("bold and italic write their complex script twin alongside", () => {
    expect(edit(null, italic(true))?.rPr).toBe("<w:rPr><w:i/><w:iCs/></w:rPr>");
  });

  it("underline writes a kind while strikethrough writes on alone", () => {
    expect(edit(null, underline(true))?.rPr).toBe(
      '<w:rPr><w:u w:val="single"/></w:rPr>'
    );
    expect(edit(null, strike(true))?.rPr).toBe("<w:rPr><w:strike/></w:rPr>");
  });

  it("a child that was not there goes into the slot OOXML's order prescribes", () => {
    const rPr =
      '<w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/>' +
      '<w:szCs w:val="20"/></w:rPr>';
    expect(edit(rPr, bold(true))?.rPr).toBe(
      '<w:rPr><w:rFonts w:ascii="Arial"/><w:b/><w:bCs/>' +
        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
    );
    expect(edit(rPr, underline(true))?.rPr).toBe(
      '<w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/>' +
        '<w:szCs w:val="20"/><w:u w:val="single"/></w:rPr>'
    );
  });

  it("formatting we do not handle stays in its original text and order", () => {
    const rPr =
      '<w:rPr><w:shd w:fill="bfbfbf" w:val="clear"/><w:rtl w:val="0"/></w:rPr>';
    expect(edit(rPr, bold(true))?.rPr).toBe(
      '<w:rPr><w:b/><w:bCs/><w:shd w:fill="bfbfbf" w:val="clear"/>' +
        '<w:rtl w:val="0"/></w:rPr>'
    );
  });

  it("writes it again with its twin even when it is already on", () => {
    expect(edit("<w:rPr><w:b/></w:rPr>", bold(true))?.rPr).toBe(
      "<w:rPr><w:b/><w:bCs/></w:rPr>"
    );
  });
});

describe("turning formatting off", () => {
  it("removes the element when nothing is inherited from a style", () => {
    const rPr = '<w:rPr><w:b/><w:bCs/><w:sz w:val="24"/></w:rPr>';
    expect(edit(rPr, bold(false))).toEqual({
      rPr: '<w:rPr><w:sz w:val="24"/></w:rPr>',
      format: { fontSizePt: 12 },
    });
  });

  it("the whole rPr disappears when no formatting is left", () => {
    expect(edit("<w:rPr><w:b/><w:bCs/></w:rPr>", bold(false))).toEqual({
      rPr: null,
      format: null,
    });
  });

  it("pins the off down when the paragraph points at a style", () => {
    expect(editStyled("<w:rPr><w:b/><w:bCs/></w:rPr>", bold(false))?.rPr).toBe(
      '<w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr>'
    );
    expect(
      editStyled('<w:rPr><w:u w:val="single"/></w:rPr>', underline(false))?.rPr
    ).toBe('<w:rPr><w:u w:val="none"/></w:rPr>');
  });

  it("pins the off down when the run points at a character style too", () => {
    const rPr = '<w:rPr><w:rStyle w:val="Strong"/><w:b/></w:rPr>';
    expect(edit(rPr, bold(false))?.rPr).toBe(
      '<w:rPr><w:rStyle w:val="Strong"/><w:b w:val="0"/>' +
        '<w:bCs w:val="0"/></w:rPr>'
    );
  });
});

describe("font size", () => {
  it("writes points as a half-point pair", () => {
    expect(edit(null, { kind: "fontSize", pt: 11 })).toEqual({
      rPr: '<w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>',
      format: { fontSizePt: 11 },
    });
    expect(edit(null, { kind: "fontSize", pt: 10.5 })?.rPr).toBe(
      '<w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
    );
  });

  it("an existing size changes in place", () => {
    const rPr =
      '<w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/>' +
      '<w:u w:val="single"/></w:rPr>';
    expect(edit(rPr, { kind: "fontSize", pt: 12 })?.rPr).toBe(
      '<w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/>' +
        '<w:u w:val="single"/></w:rPr>'
    );
  });

  it("null withdraws the setting, whether or not anything is inherited", () => {
    const rPr = '<w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>';
    expect(edit(rPr, { kind: "fontSize", pt: null })?.rPr).toBe(
      "<w:rPr><w:b/></w:rPr>"
    );
    expect(editStyled(rPr, { kind: "fontSize", pt: null })?.rPr).toBe(
      "<w:rPr><w:b/></w:rPr>"
    );
  });

  it("leaves a size that cannot be written into the document untouched", () => {
    expect(edit(null, { kind: "fontSize", pt: 0 })).toBeNull();
    expect(edit(null, { kind: "fontSize", pt: -1 })).toBeNull();
    expect(edit(null, { kind: "fontSize", pt: 10.3 })).toBeNull();
    expect(edit(null, { kind: "fontSize", pt: 2000 })).toBeNull();
    expect(edit(null, { kind: "fontSize", pt: Number.NaN })).toBeNull();
  });
});

describe("font family", () => {
  const font = (name: string | null): RunEdit => ({
    kind: "fontFamily",
    name,
  });

  it("writes a Latin name into the Latin slots alone", () => {
    expect(edit(null, font("Arial"))).toEqual({
      rPr: '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>',
      format: { fontFamily: '"Arial"' },
    });
  });

  it("writes an East Asian name into the East Asian slot as well", () => {
    // One name meant for the whole run, the way a CJK document writes it
    expect(edit(null, font("MS Mincho"))).toEqual({
      rPr:
        '<w:rPr><w:rFonts w:ascii="MS Mincho" w:hAnsi="MS Mincho" ' +
        'w:eastAsia="MS Mincho"/></w:rPr>',
      format: { fontFamily: '"MS Mincho"' },
    });
    expect(edit(null, font("Malgun Gothic"))?.rPr).toContain(
      'w:eastAsia="Malgun Gothic"'
    );
    // A name we have never heard of, written in the script itself, counts too
    expect(edit(null, font("架空明朝体"))?.rPr).toContain(
      'w:eastAsia="架空明朝体"'
    );
  });

  it("leaves the East Asian font standing when the name picked is Latin", () => {
    // A Japanese document names a Latin font beside its own on purpose, and the East
    // Asian one is the only font in the run with glyphs for its Japanese text
    const rPr =
      '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="MS Mincho" ' +
      'w:hAnsi="Times New Roman"/></w:rPr>';
    expect(edit(rPr, font("Arial"))?.rPr).toBe(
      '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" ' +
        'w:eastAsia="MS Mincho"/></w:rPr>'
    );
    // An East Asian name picked for the same run does take that slot over
    expect(edit(rPr, font("MS Gothic"))?.rPr).toBe(
      '<w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" ' +
        'w:eastAsia="MS Gothic"/></w:rPr>'
    );
  });

  it("updates the cs slot along with them when it is already there", () => {
    const rPr =
      '<w:rPr><w:rFonts w:ascii="Malgun Gothic" w:cs="Malgun Gothic" ' +
      'w:eastAsia="Malgun Gothic" w:hAnsi="Malgun Gothic"/></w:rPr>';
    expect(edit(rPr, font("Batang"))?.rPr).toBe(
      '<w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang" ' +
        'w:eastAsia="Batang" w:cs="Batang"/></w:rPr>'
    );
  });

  it("a Latin name updates the cs slot and still leaves eastAsia alone", () => {
    const rPr =
      '<w:rPr><w:rFonts w:ascii="Times New Roman" w:cs="Times New Roman" ' +
      'w:eastAsia="SimSun" w:hAnsi="Times New Roman"/></w:rPr>';
    expect(edit(rPr, font("Georgia"))?.rPr).toBe(
      '<w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" ' +
        'w:cs="Georgia" w:eastAsia="SimSun"/></w:rPr>'
    );
  });

  it("does not create cs for a run that had none", () => {
    const rPr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>';
    expect(edit(rPr, font("Georgia"))?.rPr).not.toContain("w:cs=");
  });

  it("attributes we do not set stay while a theme font is cleared away", () => {
    // A theme left in place would beat the name we wrote
    const rPr =
      '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" ' +
      'w:hint="eastAsia"/></w:rPr>';
    expect(edit(rPr, font("Dotum"))?.rPr).toBe(
      '<w:rPr><w:rFonts w:ascii="Dotum" w:hAnsi="Dotum" ' +
        'w:eastAsia="Dotum" w:hint="eastAsia"/></w:rPr>'
    );
  });

  it("a Latin name clears only the themes of the slots it writes", () => {
    const rPr =
      '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" ' +
      'w:hAnsiTheme="minorHAnsi" w:hint="eastAsia"/></w:rPr>';
    expect(edit(rPr, font("Arial"))?.rPr).toBe(
      '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" ' +
        'w:eastAsiaTheme="minorEastAsia" w:hint="eastAsia"/></w:rPr>'
    );
  });

  it("rFonts goes into the first slot OOXML prescribes", () => {
    const rPr = '<w:rPr><w:b/><w:sz w:val="20"/></w:rPr>';
    expect(edit(rPr, font("Arial"))?.rPr).toBe(
      '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>' +
        '<w:b/><w:sz w:val="20"/></w:rPr>'
    );
  });

  it("null removes the whole rFonts, whether or not anything is inherited", () => {
    const rPr = '<w:rPr><w:rFonts w:ascii="Arial" w:cs="Arial"/><w:b/></w:rPr>';
    expect(edit(rPr, font(null))?.rPr).toBe("<w:rPr><w:b/></w:rPr>");
    expect(editStyled(rPr, font(null))?.rPr).toBe("<w:rPr><w:b/></w:rPr>");
  });

  it("formatting that held nothing but the font disappears entirely", () => {
    const rPr = '<w:rPr><w:rFonts w:ascii="Arial"/></w:rPr>';
    expect(edit(rPr, font(null))).toEqual({ rPr: null, format: null });
  });

  it("leaves a name that cannot be written into the document untouched", () => {
    expect(edit(null, font(""))).toBeNull();
    expect(edit(null, font("  "))).toBeNull();
    expect(edit(null, font('Ari"al'))).toBeNull();
    expect(edit(null, font("Arial;"))).toBeNull();
  });

  it("reports that nothing needs touching when it is already that font", () => {
    expect(matchesRunEdit({ fontFamily: '"Arial"' }, font("Arial"))).toBe(true);
    expect(matchesRunEdit({ fontFamily: '"Arial"' }, font("Batang"))).toBe(
      false
    );
    expect(matchesRunEdit(null, font(null))).toBe(true);
    expect(matchesRunEdit({ fontFamily: '"Arial"' }, font(null))).toBe(false);
    // If the slots are using different names, it is not yet in the state we want
    expect(
      matchesRunEdit({ fontFamily: '"Arial","Batang"' }, font("Arial"))
    ).toBe(false);
  });
});

describe("text color and highlight", () => {
  it("writes a color as six hex digits", () => {
    expect(edit(null, { kind: "color", hex: "#2e74b5" })).toEqual({
      rPr: '<w:rPr><w:color w:val="2E74B5"/></w:rPr>',
      format: { color: "#2E74B5" },
    });
    expect(edit(null, { kind: "color", hex: "FF0000" })?.rPr).toBe(
      '<w:rPr><w:color w:val="FF0000"/></w:rPr>'
    );
  });

  it("leaves a value that is not a color untouched", () => {
    expect(edit(null, { kind: "color", hex: "red" })).toBeNull();
    expect(edit(null, { kind: "color", hex: "#12345" })).toBeNull();
  });

  it("turning the color off pins auto down only when something is inherited", () => {
    const rPr = '<w:rPr><w:color w:val="FF0000"/><w:b/></w:rPr>';
    expect(edit(rPr, { kind: "color", hex: null })?.rPr).toBe(
      "<w:rPr><w:b/></w:rPr>"
    );
    expect(editStyled(rPr, { kind: "color", hex: null })?.rPr).toBe(
      '<w:rPr><w:color w:val="auto"/><w:b/></w:rPr>'
    );
  });

  it("writes a background color as w:shd", () => {
    expect(edit(null, { kind: "background", hex: "#FFF2CC" })).toEqual({
      rPr: '<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:rPr>',
      format: { background: "#FFF2CC" },
    });
  });

  it("painting a background color clears the old highlight away with it", () => {
    // Word paints the highlight on top of the shading. Left in place, the new color would not be visible
    expect(
      edit('<w:rPr><w:highlight w:val="cyan"/></w:rPr>', {
        kind: "background",
        hex: "#FF0000",
      })
    ).toEqual({
      rPr: '<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:rPr>',
      format: { background: "#FF0000" },
    });
  });

  it("turning the background off removes the shading and the highlight together", () => {
    const rPr =
      '<w:rPr><w:highlight w:val="cyan"/>' +
      '<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:rPr>';
    expect(edit(rPr, { kind: "background", hex: null })).toEqual({
      rPr: null,
      format: null,
    });
  });

  it("pins the background off down when something is inherited", () => {
    const rPr = '<w:rPr><w:b/><w:shd w:val="clear" w:fill="FF0000"/></w:rPr>';
    expect(editStyled(rPr, { kind: "background", hex: null })?.rPr).toBe(
      '<w:rPr><w:b/><w:highlight w:val="none"/>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:rPr>'
    );
  });

  it("a value that is not a color leaves the original untouched", () => {
    expect(edit(null, { kind: "background", hex: "red" })).toBeNull();
  });

  it("fixing the background color leaves the run's other formatting as it is", () => {
    const rPr = '<w:rPr><w:b/><w:sz w:val="24"/><w:vanish/></w:rPr>';
    expect(edit(rPr, { kind: "background", hex: "#FFFF00" })?.rPr).toBe(
      '<w:rPr><w:b/><w:sz w:val="24"/><w:vanish/>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/></w:rPr>'
    );
  });
});

describe("deriving the display values again", () => {
  it("the same rPr always becomes the same display values", () => {
    const rPr =
      '<w:rPr><w:b/><w:sz w:val="24"/><w:highlight w:val="cyan"/></w:rPr>';
    expect(readRunProps(rPr)).toEqual({
      bold: true,
      fontSizePt: 12,
      highlight: "cyan",
    });
    expect(edit(rPr, italic(true))?.format).toEqual({
      bold: true,
      italic: true,
      fontSizePt: 12,
      highlight: "cyan",
    });
  });

  it("display values inherited from a style survive the operation", () => {
    const next = editRunProps(
      // The color the paragraph style gave is not in the rPr, only in the display values
      {
        rPr: "<w:rPr><w:b/></w:rPr>",
        format: { bold: true, color: "#2E74B5" },
      },
      '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>',
      bold(false)
    );
    expect(next).toEqual({
      rPr: '<w:rPr><w:b w:val="0"/><w:bCs w:val="0"/></w:rPr>',
      format: { color: "#2E74B5" },
    });
  });

  it("leaves an rPr whose shape it does not recognize untouched", () => {
    expect(edit("<w:rPr><w:b/>", bold(true))).toBeNull();
    expect(edit("not formatting", bold(true))).toBeNull();
  });
});

describe("checking whether it is already in the state we want", () => {
  it("reads on/off formatting from the display values", () => {
    expect(isRunToggleOn({ bold: true }, "bold")).toBe(true);
    expect(isRunToggleOn({ bold: true }, "italic")).toBe(false);
    expect(isRunToggleOn(null, "bold")).toBe(false);
    // An underline holding a kind means it is on
    expect(isRunToggleOn({ underline: "double" }, "underline")).toBe(true);
  });

  it("reports that nothing needs touching when it is already in that state", () => {
    expect(matchesRunEdit({ bold: true }, bold(true))).toBe(true);
    expect(matchesRunEdit(null, bold(false))).toBe(true);
    expect(
      matchesRunEdit({ fontSizePt: 11 }, { kind: "fontSize", pt: 11 })
    ).toBe(true);
    expect(matchesRunEdit(null, { kind: "fontSize", pt: null })).toBe(true);
    // Colors differing only in capitalization are the same color
    expect(
      matchesRunEdit({ color: "#2e74b5" }, { kind: "color", hex: "#2E74B5" })
    ).toBe(true);
    expect(
      matchesRunEdit(
        { background: "#ffff00" },
        { kind: "background", hex: "#FFFF00" }
      )
    ).toBe(true);
    // A highlight still in place means the move over to shading has not happened yet
    expect(
      matchesRunEdit(
        { highlight: "yellow" },
        { kind: "background", hex: "#FFFF00" }
      )
    ).toBe(false);
    expect(matchesRunEdit(null, { kind: "background", hex: null })).toBe(true);
    expect(
      matchesRunEdit({ highlight: "yellow" }, { kind: "background", hex: null })
    ).toBe(false);
  });
});
