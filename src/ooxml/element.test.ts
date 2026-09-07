// @vitest-environment node
import { describe, expect, it } from "vitest";
import {
  elementXml,
  emptyTagXml,
  openTagXml,
  wAttrValue,
  withAttr,
  withoutAttrs,
  wLocalName,
  type XmlAttr,
  xmlnsAttr,
} from "./element";
import { wName } from "./names";

describe("writing an element", () => {
  /**
   * A style id is one the document names, so it may hold anything a name may hold. Written into
   * the value as it stands, an `&` would leave a part Word refuses to open.
   */
  it("escapes every attribute value", () => {
    expect(
      elementXml(wName("pStyle"), [[wName("val"), 'Heading & <"1">']])
    ).toBe('<w:pStyle w:val="Heading &amp; &lt;&quot;1&quot;&gt;"/>');
  });

  it("writes a self-closing tag when there are no children", () => {
    expect(elementXml(wName("b"), [])).toBe("<w:b/>");
    expect(elementXml(wName("sz"), [[wName("val"), "24"]])).toBe(
      '<w:sz w:val="24"/>'
    );
    // The empty text is a child, so a caller that wants an empty pair of tags gets one
    expect(elementXml(wName("t"), [["xml:space", "preserve"]], [""])).toBe(
      '<w:t xml:space="preserve"></w:t>'
    );
    expect(elementXml(wName("r"), [], ["<w:br/>"])).toBe("<w:r><w:br/></w:r>");
  });

  it("keeps the slot of an attribute it rewrites and appends a new one", () => {
    const ind: XmlAttr[] = [
      ["w:left", "720"],
      ["w:right", "200"],
    ];
    expect(elementXml(wName("ind"), withAttr(ind, "left", "1440"))).toBe(
      '<w:ind w:left="1440" w:right="200"/>'
    );
    expect(elementXml(wName("ind"), withAttr(ind, "firstLine", "360"))).toBe(
      '<w:ind w:left="720" w:right="200" w:firstLine="360"/>'
    );
  });

  /**
   * A producer may qualify an attribute of its own with a local name WordprocessingML also uses.
   * `mc:Ignorable` tells a consumer to skip it, so reading it as the formatting value, or writing
   * the formatting value into it, would take an attribute nobody reads for the one Word does.
   */
  it("reads the WordprocessingML attribute past a foreign one of the same local name", () => {
    const top: XmlAttr[] = [
      ["x:val", "none"],
      ["w:val", "single"],
    ];
    expect(wAttrValue(top, "val")).toBe("single");
    expect(wAttrValue([["x:val", "none"]], "val")).toBeNull();
    expect(wAttrValue([], "val")).toBeNull();
  });

  it("reads an unprefixed attribute as the WordprocessingML one when no w: spelling stands", () => {
    expect(wAttrValue([["val", "single"]], "val")).toBe("single");
    expect(
      wAttrValue(
        [
          ["val", "none"],
          ["w:val", "single"],
        ],
        "val"
      )
    ).toBe("single");
    expect(wLocalName("w:val")).toBe("val");
    expect(wLocalName("val")).toBe("val");
    expect(wLocalName("x:val")).toBeNull();
    expect(wLocalName("xmlns:w")).toBeNull();
  });

  it("writes and removes the WordprocessingML attribute and leaves a foreign one standing", () => {
    const top: XmlAttr[] = [
      ["x:val", "none"],
      ["w:val", "single"],
    ];
    expect(elementXml(wName("top"), withAttr(top, "val", "none"))).toBe(
      '<w:top x:val="none" w:val="none"/>'
    );
    expect(elementXml(wName("top"), withAttr(top, "val", null))).toBe(
      '<w:top x:val="none"/>'
    );
    expect(elementXml(wName("top"), withoutAttrs(top, ["val"]))).toBe(
      '<w:top x:val="none"/>'
    );
    // With no WordprocessingML spelling there yet, the value is added rather than written into the foreign one
    expect(
      elementXml(
        wName("shd"),
        withAttr([["x:fill", "FF0000"]], "fill", "00FF00")
      )
    ).toBe('<w:shd x:fill="FF0000" w:fill="00FF00"/>');
  });

  it("removes only the named local names under the prefix and keeps the rest", () => {
    const paragraph: XmlAttr[] = [
      ["w14:paraId", "1F2A"],
      ["w:rsidR", "00A1"],
      ["w14:textId", "3B4C"],
    ];
    // `paraId` is written under `w14`, so the WordprocessingML reading leaves it where it stands
    expect(elementXml(wName("p"), withoutAttrs(paragraph, ["paraId"]))).toBe(
      '<w:p w14:paraId="1F2A" w:rsidR="00A1" w14:textId="3B4C"/>'
    );
    expect(
      elementXml(wName("p"), withoutAttrs(paragraph, ["paraId"], "w14"))
    ).toBe('<w:p w:rsidR="00A1" w14:textId="3B4C"/>');
    // A name neither vocabulary was asked for keeps every attribute as it was written
    expect(
      elementXml(wName("p"), withoutAttrs(paragraph, ["rsidR"], "w14"))
    ).toBe('<w:p w14:paraId="1F2A" w:rsidR="00A1" w14:textId="3B4C"/>');
  });

  it("removes an attribute when the value is null", () => {
    const spacing: XmlAttr[] = [
      ["w:before", "120"],
      ["w:after", "240"],
    ];
    expect(
      elementXml(wName("spacing"), withAttr(spacing, "before", null))
    ).toBe('<w:spacing w:after="240"/>');
    expect(
      elementXml(wName("spacing"), withoutAttrs(spacing, ["before", "after"]))
    ).toBe("<w:spacing/>");
  });

  it("splices the raw attribute text of an opening tag verbatim", () => {
    // Whatever a producer wrote stands as it stands: the order, the spelling, the escaping
    const raw = 'w14:paraId="1F2A" w:rsidR="00A1"  w:rsidRDefault="00A1"';
    expect(openTagXml(wName("p"), raw)).toBe(`<w:p ${raw}>`);
    expect(openTagXml(wName("p"), null)).toBe("<w:p>");
    expect(emptyTagXml(wName("br"), 'w:type="page"')).toBe(
      '<w:br w:type="page"/>'
    );
    expect(emptyTagXml(wName("br"), null)).toBe("<w:br/>");
  });

  it("declares a prefix the part a fragment lands in may not know", () => {
    expect(
      elementXml(
        "w15:person",
        [xmlnsAttr("w15"), ["w15:author", "Ada & Co"]],
        []
      )
    ).toBe(
      '<w15:person xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"' +
        ' w15:author="Ada &amp; Co"/>'
    );
  });
});
