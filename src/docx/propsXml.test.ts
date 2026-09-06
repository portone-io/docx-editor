// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { childByLocalName } from "../ooxml/xml";
import { readRunFormat } from "./formatting";
import {
  innerXml,
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
  const propsOf = (xml: string) => {
    const props = parseProps(xml);
    if (!props) throw new Error("could not read the fragment");
    return props;
  };

  it("an existing child changes in place", () => {
    const next = setPropsChild(
      propsOf(
        '<w:tcPr><w:tcW w:w="1"/><w:gridSpan w:val="2"/>' +
          '<w:vAlign w:val="center"/></w:tcPr>'
      ),
      "gridSpan",
      '<w:gridSpan w:val="3"/>',
      TC_PR_ORDER
    );
    expect(next.children.map((child) => child.xml)).toEqual([
      '<w:tcW w:w="1"/>',
      '<w:gridSpan w:val="3"/>',
      '<w:vAlign w:val="center"/>',
    ]);
  });

  it("removes the child when it is null", () => {
    const next = setPropsChild(
      propsOf('<w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr>'),
      "gridSpan",
      null,
      TC_PR_ORDER
    );
    expect(next.children.map((child) => child.name)).toEqual(["vMerge"]);
  });

  it("a child that was not there goes into the slot OOXML's order prescribes", () => {
    const next = setPropsChild(
      propsOf(
        '<w:tcPr><w:tcW w:w="1"/><w:tcBorders/><w:vAlign w:val="center"/></w:tcPr>'
      ),
      "vMerge",
      "<w:vMerge/>",
      TC_PR_ORDER
    );
    expect(next.children.map((child) => child.name)).toEqual([
      "tcW",
      "vMerge",
      "tcBorders",
      "vAlign",
    ]);
  });

  it("a child whose order is unknown stays behind the child ahead of it", () => {
    const next = setPropsChild(
      propsOf("<w:tcPr><w:tcW/><w:unknownThing/><w:vAlign/></w:tcPr>"),
      "vMerge",
      "<w:vMerge/>",
      TC_PR_ORDER
    );
    expect(next.children.map((child) => child.name)).toEqual([
      "tcW",
      "unknownThing",
      "vMerge",
      "vAlign",
    ]);
  });

  it("a changed child keeps what stood in front of it", () => {
    const next = setPropsChild(
      propsOf('<w:tcPr>\n  <w:gridSpan w:val="2"/></w:tcPr>'),
      "gridSpan",
      '<w:gridSpan w:val="3"/>',
      TC_PR_ORDER
    );
    expect(renderProps(next)).toBe(
      '<w:tcPr>\n  <w:gridSpan w:val="3"/></w:tcPr>'
    );
  });

  it("a removed child leaves what stood in front of it to the child that follows", () => {
    const next = setPropsChild(
      propsOf("<w:tcPr><!-- kept --><w:gridSpan/><w:vMerge/></w:tcPr>"),
      "gridSpan",
      null,
      TC_PR_ORDER
    );
    expect(renderProps(next)).toBe("<w:tcPr><!-- kept --><w:vMerge/></w:tcPr>");
  });

  it("a removed last child leaves it to the tail", () => {
    const next = setPropsChild(
      propsOf("<w:tcPr><w:vMerge/><!-- kept --><w:gridSpan/></w:tcPr>"),
      "gridSpan",
      null,
      TC_PR_ORDER
    );
    expect(renderProps(next)).toBe("<w:tcPr><w:vMerge/><!-- kept --></w:tcPr>");
  });
});

describe("innerXml", () => {
  it.each([
    ['<w:tcW w:w="1"/>', ""],
    ['<w:tcW w:w="1"></w:tcW>', ""],
    ["<w:tcW><!-- why --></w:tcW>", "<!-- why -->"],
    ["<w:tcW><!-- </x> --></w:tcW>", "<!-- </x> -->"],
    // The opening tag is read rather than scanned for, so this `>` is not the end of it
    ['<w:tcW w:w="1" w:note="a>b">text</w:tcW>', "text"],
    ["<w:tcW><w:tcW/></w:tcW>", "<w:tcW/>"],
    ["not an element", ""],
  ])("reads what stood inside %s", (xml, inner) => {
    expect(innerXml(xml)).toBe(inner);
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

  it.each([
    '<w:tcPr>\n  <w:tcW w:w="1"/>\n  <w:vMerge/>\n</w:tcPr>',
    "<w:tcPr><!-- why --><w:vMerge/></w:tcPr>",
    "<w:tcPr><w:vMerge/><?spell off?></w:tcPr>",
    "<w:tcPr><![CDATA[note]]><w:vMerge/></w:tcPr>",
  ])("writes %s back character for character", (xml) => {
    const props = parseProps(xml);
    if (!props) throw new Error("could not read the fragment");
    expect(renderProps(props)).toBe(xml);
  });

  it("writes a fragment whose only content is a comment, and none for whitespace alone", () => {
    const comment = parseProps("<w:tcPr><!-- why --></w:tcPr>");
    const blank = parseProps("<w:tcPr>\n</w:tcPr>");
    if (!comment || !blank) throw new Error("could not read the fragment");
    expect(renderProps(comment)).toBe("<w:tcPr><!-- why --></w:tcPr>");
    expect(renderProps(blank)).toBe("");
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
