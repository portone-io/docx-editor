// @vitest-environment node
import { describe, expect, it } from "vitest";
import {
  attrValue,
  elementXml,
  emptyTagXml,
  openTagXml,
  withAttr,
  withoutAttrs,
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

  it("matches an attribute by its local part so a foreign prefix survives", () => {
    const shd: XmlAttr[] = [["x:fill", "FF0000"]];
    expect(attrValue(shd, "fill")).toBe("FF0000");
    expect(elementXml("x:shd", withAttr(shd, "fill", "00FF00"))).toBe(
      '<x:shd x:fill="00FF00"/>'
    );
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
