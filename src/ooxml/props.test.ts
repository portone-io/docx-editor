// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { readRunFormat } from "../docx/formatting";
import {
  attrsOf,
  childElement,
  editChild,
  innerXml,
  parseProps,
  parsePropsXml,
  propsChild,
  renderElement,
  renderProps,
  setChild,
  withAttrs,
} from "./props";
import { childByLocalName } from "./xml";

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

describe("setChild", () => {
  const propsOf = (xml: string) => {
    const props = parseProps(xml);
    if (!props) throw new Error("could not read the fragment");
    return props;
  };

  it("an existing child changes in place", () => {
    const next = setChild(
      propsOf(
        '<w:tcPr><w:tcW w:w="1"/><w:gridSpan w:val="2"/>' +
          '<w:vAlign w:val="center"/></w:tcPr>'
      ),
      "gridSpan",
      '<w:gridSpan w:val="3"/>'
    );
    expect(next.children.map((child) => child.xml)).toEqual([
      '<w:tcW w:w="1"/>',
      '<w:gridSpan w:val="3"/>',
      '<w:vAlign w:val="center"/>',
    ]);
  });

  it("removes the child when it is null", () => {
    const next = setChild(
      propsOf('<w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr>'),
      "gridSpan",
      null
    );
    expect(next.children.map((child) => child.name)).toEqual(["vMerge"]);
  });

  it("a child that was not there goes into the slot OOXML's order prescribes", () => {
    const next = setChild(
      propsOf(
        '<w:tcPr><w:tcW w:w="1"/><w:tcBorders/><w:vAlign w:val="center"/></w:tcPr>'
      ),
      "vMerge",
      "<w:vMerge/>"
    );
    expect(next.children.map((child) => child.name)).toEqual([
      "tcW",
      "vMerge",
      "tcBorders",
      "vAlign",
    ]);
  });

  it("a child whose order is unknown stays behind the child ahead of it", () => {
    const next = setChild(
      propsOf("<w:tcPr><w:tcW/><w:unknownThing/><w:vAlign/></w:tcPr>"),
      "vMerge",
      "<w:vMerge/>"
    );
    expect(next.children.map((child) => child.name)).toEqual([
      "tcW",
      "unknownThing",
      "vMerge",
      "vAlign",
    ]);
  });

  it("a changed child keeps what stood in front of it", () => {
    const next = setChild(
      propsOf('<w:tcPr>\n  <w:gridSpan w:val="2"/></w:tcPr>'),
      "gridSpan",
      '<w:gridSpan w:val="3"/>'
    );
    expect(renderProps(next)).toBe(
      '<w:tcPr>\n  <w:gridSpan w:val="3"/></w:tcPr>'
    );
  });

  it("a removed child leaves what stood in front of it to the child that follows", () => {
    const next = setChild(
      propsOf("<w:tcPr><!-- kept --><w:gridSpan/><w:vMerge/></w:tcPr>"),
      "gridSpan",
      null
    );
    expect(renderProps(next)).toBe("<w:tcPr><!-- kept --><w:vMerge/></w:tcPr>");
  });

  it("a removed last child leaves it to the tail", () => {
    const next = setChild(
      propsOf("<w:tcPr><w:vMerge/><!-- kept --><w:gridSpan/></w:tcPr>"),
      "gridSpan",
      null
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

describe("attrsOf and withAttrs", () => {
  it("keeps the raw attribute text of an untouched fragment", () => {
    const xml = "<w:tcPr w:x='1'\n  w:y=\"a&gt;b\"><w:vMerge/></w:tcPr>";
    const props = parseProps(xml);
    if (!props) throw new Error("could not read the fragment");
    expect(attrsOf(props)).toEqual([
      ["w:x", "1"],
      ["w:y", "a>b"],
    ]);
    expect(renderProps(props)).toBe(xml);
  });

  it("writes the pairs again only where they were edited", () => {
    const props = parseProps('<w:ind w:left="720" w:right="60"/>');
    if (!props) throw new Error("could not read the fragment");
    expect(renderElement(withAttrs(props, [["w:left", "0"]]))).toBe(
      '<w:ind w:left="0"/>'
    );
    expect(renderElement(withAttrs(props, []))).toBe("<w:ind/>");
  });

  it("is null for an attribute string a parser would refuse", () => {
    const props = parseProps("<w:tcPr w:x=1><w:vMerge/></w:tcPr>");
    expect(props && attrsOf(props)).toBeNull();
  });
});

describe("childElement", () => {
  it("reads the tag and the pairs of one child, and says when a child is not there", () => {
    const props = parseProps('<w:tcPr><w14:shd w:fill="FF0000"/></w:tcPr>');
    if (!props) throw new Error("could not read the fragment");
    expect(childElement(props, "shd")).toEqual({
      tag: "w14:shd",
      attrs: [["w:fill", "FF0000"]],
    });
    expect(childElement(props, "vAlign")).toEqual({ tag: null, attrs: [] });
  });

  it("is null for a child whose shape cannot be made out", () => {
    const props = parseProps("<w:tcPr><w:shd w:fill=red/></w:tcPr>");
    expect(props && childElement(props, "shd")).toBeNull();
  });
});

describe("renderElement", () => {
  it("renders an empty leaf as a self-closing element", () => {
    expect(renderElement({ tag: "w:sdtPr", attrs: null, children: [] })).toBe(
      "<w:sdtPr/>"
    );
    expect(
      renderElement({ tag: "w:sdtPr", attrs: 'w:x="1"', children: [] })
    ).toBe('<w:sdtPr w:x="1"/>');
  });

  it("writes the children of one that holds something", () => {
    const props = parseProps("<w:tcMar><w:top/></w:tcMar>");
    if (!props) throw new Error("could not read the fragment");
    expect(renderElement(props)).toBe("<w:tcMar><w:top/></w:tcMar>");
  });
});

describe("editChild", () => {
  const tcPr = (children: string) => `<w:tcPr>${children}</w:tcPr>`;
  const propsOf = (xml: string) => {
    const props = parseProps(xml);
    if (!props) throw new Error("could not read the fragment");
    return props;
  };

  it("edits a nested child along its path and leaves the siblings byte for byte", () => {
    const original = tcPr(
      '<w:tcW w:w="1"/>\n  <w:tcBorders><!-- why --><w:top w:val="single"/>' +
        '<w:bottom w:val="none"/></w:tcBorders><w:shd w:fill="FF0000"/>'
    );
    const edited = editChild(propsOf(original), ["tcBorders", "top"], () => ({
      tag: "w:top",
      attrs: 'w:val="double"',
      children: [],
    }));
    expect(edited && renderProps(edited)).toBe(
      original.replace('<w:top w:val="single"/>', '<w:top w:val="double"/>')
    );
  });

  it("writes a container that was not there and takes an emptied one away", () => {
    const written = editChild(
      propsOf(tcPr('<w:tcW w:w="1"/>')),
      ["tcMar", "top"],
      () => ({ tag: "w:top", attrs: 'w:w="80"', children: [] })
    );
    expect(written && renderProps(written)).toBe(
      tcPr('<w:tcW w:w="1"/><w:tcMar><w:top w:w="80"/></w:tcMar>')
    );
    const emptied = editChild(
      propsOf(tcPr('<w:tcMar><w:top w:w="80"/></w:tcMar><w:vMerge/>')),
      ["tcMar", "top"],
      () => null
    );
    expect(emptied && renderProps(emptied)).toBe(tcPr("<w:vMerge/>"));
  });

  it("puts a nested child that was not there into the spot the order calls for", () => {
    const edited = editChild(
      propsOf(tcPr('<w:tcBorders><w:bottom w:val="single"/></w:tcBorders>')),
      ["tcBorders", "top"],
      () => ({ tag: "w:top", attrs: null, children: [] })
    );
    expect(edited && renderProps(edited)).toBe(
      tcPr('<w:tcBorders><w:top/><w:bottom w:val="single"/></w:tcBorders>')
    );
  });

  it("is null when a fragment on the path cannot be made out", () => {
    const broken = {
      tag: "w:tcPr",
      attrs: null,
      children: [{ name: "tcBorders", xml: "<w:tcBorders><w:top/>" }],
    };
    expect(editChild(broken, ["tcBorders", "top"], () => null)).toBeNull();
    expect(editChild(broken, ["tcBorders"], () => null)).toBeNull();
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
