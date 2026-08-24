// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { childByLocalName } from "../ooxml/xml";
import { readRunFormat } from "./formatting";
import {
  P_PR_ORDER,
  parseProps,
  parsePropsXml,
  propsChild,
  RUN_PR_ORDER,
  renderProps,
  setPropsChild,
  TC_PR_ORDER,
} from "./propsXml";

/** Pulls out just the child names, to keep the tests easy to read */
function names(xml: string): string[] {
  const props = parseProps(xml);
  if (!props) throw new Error("could not read the fragment");
  return props.children.map((child) => child.name);
}

describe("parseProps", () => {
  it("splits the children into a name and the original fragment", () => {
    const props = parseProps(
      '<w:tcPr><w:tcW w:type="dxa" w:w="1000"/><w:vMerge/></w:tcPr>'
    );
    expect(props).toEqual({
      tag: "w:tcPr",
      attrs: null,
      children: [
        { name: "tcW", xml: '<w:tcW w:type="dxa" w:w="1000"/>' },
        { name: "vMerge", xml: "<w:vMerge/>" },
      ],
    });
  });

  it("a child holding children of its own is one whole fragment", () => {
    expect(
      names(
        "<w:tcPr><w:tcBorders><w:top/><w:bottom/></w:tcBorders>" +
          '<w:vAlign w:val="center"/></w:tcPr>'
      )
    ).toEqual(["tcBorders", "vAlign"]);
    const props = parseProps(
      "<w:tcPr><w:tcBorders><w:top/></w:tcBorders></w:tcPr>"
    );
    expect(props?.children[0].xml).toBe("<w:tcBorders><w:top/></w:tcBorders>");
  });

  it("reads a fragment with no children and the opening tag's attributes", () => {
    expect(parseProps("<w:tcPr/>")).toEqual({
      tag: "w:tcPr",
      attrs: null,
      children: [],
    });
    expect(
      parseProps('<w:tblPr w:x="1"><w:jc w:val="left"/></w:tblPr>')
    ).toEqual({
      tag: "w:tblPr",
      attrs: 'w:x="1"',
      children: [{ name: "jc", xml: '<w:jc w:val="left"/>' }],
    });
  });

  it("does not read an angle bracket inside an attribute value as the end of the tag", () => {
    expect(names('<w:tcPr><w:tcW w:w="a&gt;b"/></w:tcPr>')).toEqual(["tcW"]);
  });

  it("is null when it does not recognize the shape", () => {
    expect(parseProps("<w:tcPr><w:tcW/>")).toBeNull();
    expect(parseProps("<w:tcPr/><w:extra/>")).toBeNull();
    expect(parseProps("text only")).toBeNull();
  });
});

describe("setPropsChild", () => {
  const children = (xml: string) => parseProps(xml)?.children ?? [];

  it("an existing child changes in place", () => {
    const next = setPropsChild(
      children(
        '<w:tcPr><w:tcW w:w="1"/><w:gridSpan w:val="2"/>' +
          '<w:vAlign w:val="center"/></w:tcPr>'
      ),
      "gridSpan",
      '<w:gridSpan w:val="3"/>',
      TC_PR_ORDER
    );
    expect(next.map((child) => child.xml)).toEqual([
      '<w:tcW w:w="1"/>',
      '<w:gridSpan w:val="3"/>',
      '<w:vAlign w:val="center"/>',
    ]);
  });

  it("removes the child when it is null", () => {
    const next = setPropsChild(
      children('<w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr>'),
      "gridSpan",
      null,
      TC_PR_ORDER
    );
    expect(next.map((child) => child.name)).toEqual(["vMerge"]);
  });

  it("a child that was not there goes into the slot OOXML's order prescribes", () => {
    const next = setPropsChild(
      children(
        '<w:tcPr><w:tcW w:w="1"/><w:tcBorders/><w:vAlign w:val="center"/></w:tcPr>'
      ),
      "vMerge",
      "<w:vMerge/>",
      TC_PR_ORDER
    );
    expect(next.map((child) => child.name)).toEqual([
      "tcW",
      "vMerge",
      "tcBorders",
      "vAlign",
    ]);
  });

  it("a child whose order is unknown stays behind the child ahead of it", () => {
    const next = setPropsChild(
      children("<w:tcPr><w:tcW/><w:unknownThing/><w:vAlign/></w:tcPr>"),
      "vMerge",
      "<w:vMerge/>",
      TC_PR_ORDER
    );
    expect(next.map((child) => child.name)).toEqual([
      "tcW",
      "unknownThing",
      "vMerge",
      "vAlign",
    ]);
  });
});

describe("renderProps", () => {
  it("writes no fragment when there are no children", () => {
    expect(renderProps({ tag: "w:tcPr", attrs: null, children: [] })).toBe("");
  });

  it("joins the child fragments character for character as they were", () => {
    const xml = '<w:tcPr><w:tcW w:type="dxa" w:w="1000"/><w:vMerge/></w:tcPr>';
    const props = parseProps(xml);
    if (!props) throw new Error("could not read the fragment");
    expect(renderProps(props)).toBe(xml);
    expect(propsChild(props.children, "vMerge")?.xml).toBe("<w:vMerge/>");
  });
});

describe("the child order tables", () => {
  it("each of the three order tables lists every name exactly once", () => {
    for (const order of [RUN_PR_ORDER, P_PR_ORDER, TC_PR_ORDER]) {
      expect(new Set(order).size).toBe(order.length);
    }
  });
});

describe("parsePropsXml", () => {
  it("reads back a fragment that was kept aside without its declarations", () => {
    const el = parsePropsXml('<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>');
    expect(readRunFormat(el)).toEqual({ bold: true, fontSizePt: 12 });
  });

  it("reads it even with a prefix we do not read mixed in", () => {
    const el = parsePropsXml(
      '<w:rPr><w14:ligatures w14:val="none"/><w:b/></w:rPr>'
    );
    expect(readRunFormat(el)).toEqual({ bold: true });
  });

  it("reads it even with a reserved prefix that cannot be redeclared mixed in", () => {
    // Redeclaring `xml:` makes parsing fail, which turns into a silent no-op
    const el = parsePropsXml(
      '<w:pPr><w:rPr><w:t xml:space="preserve"/></w:rPr><w:jc w:val="both"/></w:pPr>'
    );
    expect(el && childByLocalName(el, "jc")).not.toBeNull();
  });

  it("is null when it does not recognize the shape", () => {
    expect(parsePropsXml("<w:rPr><w:b/>")).toBeNull();
    expect(parsePropsXml("")).toBeNull();
  });
});
