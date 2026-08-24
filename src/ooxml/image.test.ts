// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { anchoredDrawingXml, inlineDrawingXml } from "../__testing__/docx";
import {
  emuToPx,
  imageDrawingXml,
  pxToEmu,
  readDrawingPicture,
  toImageExtent,
  toImageSrc,
  withExtent,
} from "./image";
import { parseXml } from "./xml";

const W_NS_DECL =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** A drawing on its own, with the prefix the document element would have declared for it */
function drawing(xml: string): Element {
  return parseXml(xml.replace("<w:drawing", `<w:drawing ${W_NS_DECL}`))
    .documentElement;
}

describe("reading a drawing", () => {
  it("reads an inline picture", () => {
    const picture = readDrawingPicture(
      drawing(inlineDrawingXml({ relId: "rId9", descr: "Seal" }))
    );
    expect(picture).toEqual({
      relId: "rId9",
      extent: { cx: 1905000, cy: 952500 },
      alt: "Seal",
    });
  });

  it("reads no alternative text when there is none", () => {
    expect(readDrawingPicture(drawing(inlineDrawingXml()))?.alt).toBe(null);
  });

  /** Everything below stays on the preservation path, and the paragraph goes out untouched */
  it("does not read a floating picture", () => {
    expect(readDrawingPicture(drawing(anchoredDrawingXml()))).toBe(null);
  });

  it("does not read a graphic that is not a picture", () => {
    const chart = inlineDrawingXml({
      uri: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    });
    expect(readDrawingPicture(drawing(chart))).toBe(null);
  });

  it("does not read a picture linked to a file outside the package", () => {
    expect(
      readDrawingPicture(drawing(inlineDrawingXml({ linked: true })))
    ).toBe(null);
  });

  it("does not read a picture with no size to draw it at", () => {
    expect(readDrawingPicture(drawing(inlineDrawingXml({ cx: 0 })))).toBe(null);
  });
});

describe("the size a drawing records", () => {
  it("is one pixel every 9525 units", () => {
    expect(emuToPx(1905000)).toBe(200);
    expect(pxToEmu(200)).toBe(1905000);
  });

  it("accepts only whole positive numbers", () => {
    expect(toImageExtent({ cx: 100, cy: 50 })).toEqual({ cx: 100, cy: 50 });
    expect(toImageExtent({ cx: 100, cy: 0 })).toBe(null);
    expect(toImageExtent({ cx: 1.5, cy: 50 })).toBe(null);
    expect(toImageExtent({ cx: "100", cy: "50" })).toEqual({ cx: 100, cy: 50 });
    expect(toImageExtent({ cx: -100, cy: 50 })).toBe(null);
    expect(toImageExtent(null)).toBe(null);
  });
});

/** The src leaves for the screen as an img attribute and comes back in again */
describe("the src an image carries", () => {
  it("accepts only a data URL of a kind a browser draws", () => {
    const png = "data:image/png;base64,iVBORw0KGgo=";
    expect(toImageSrc(png)).toBe(png);
    expect(toImageSrc("data:image/webp;base64,AAAA")).toBe(null);
    expect(toImageSrc("data:image/svg+xml;base64,AAAA")).toBe(null);
  });

  it("lets nothing else through", () => {
    expect(toImageSrc("https://example.com/seal.png")).toBe(null);
    expect(toImageSrc("javascript:alert(1)")).toBe(null);
    expect(toImageSrc('data:image/png;base64,AAA" onerror="alert(1)')).toBe(
      null
    );
    expect(toImageSrc(null)).toBe(null);
  });
});

describe("rewriting the size in a drawing", () => {
  const xml = inlineDrawingXml({ cx: 1905000, cy: 952500 });

  it("changes both extents and nothing else", () => {
    const resized = withExtent(xml, { cx: 952500, cy: 476250 });
    expect(resized).toContain('<wp:extent cx="952500" cy="476250"/>');
    expect(resized).toContain('<a:ext cx="952500" cy="476250"/>');
    expect(resized).not.toContain("1905000");
    // The wrapping effect extent is a different element and is left alone
    expect(resized).toContain('<wp:effectExtent l="0" t="0" r="0" b="0"/>');
    expect(resized.replace(/cx="\d+" cy="\d+"/g, "")).toBe(
      xml.replace(/cx="\d+" cy="\d+"/g, "")
    );
  });

  it("gives back the very same string for the size it came in with", () => {
    expect(withExtent(xml, { cx: 1905000, cy: 952500 })).toBe(xml);
  });
});

describe("the drawing built for an inserted image", () => {
  const xml = imageDrawingXml({
    relId: "rId3",
    docPrId: 12,
    extent: { cx: 952500, cy: 476250 },
    alt: 'a quoted " description',
  });

  it("reads back as the picture it was built from", () => {
    expect(readDrawingPicture(drawing(xml))).toEqual({
      relId: "rId3",
      extent: { cx: 952500, cy: 476250 },
      alt: 'a quoted " description',
    });
  });

  it("states the size in both places Word looks", () => {
    expect(xml).toContain('<wp:extent cx="952500" cy="476250"/>');
    expect(xml).toContain('<a:ext cx="952500" cy="476250"/>');
  });

  it("escapes the alternative text into the attribute", () => {
    expect(xml).toContain('descr="a quoted &quot; description"');
  });
});
