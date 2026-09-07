import { describe, expect, it } from "vitest";
import { DocxExportError } from "./errors";
import { NAMESPACES, xmlnsDecl } from "./names";
import {
  ensureRootDeclarations,
  partRootProblem,
  type RootDeclarations,
  splicePart,
} from "./partSplice";

const PROLOG = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** The code and message an export refusal carries */
function refusal(run: () => unknown): { code: string; message: string } {
  try {
    run();
  } catch (error) {
    if (error instanceof DocxExportError) {
      return { code: error.code, message: error.message };
    }
    throw error;
  }
  throw new Error("the splice was expected to be refused");
}

describe("splicePart", () => {
  it("appends before the closing tag of a prefixed root", () => {
    const xml =
      `${PROLOG}<pr:Relationships xmlns:pr="${RELS_NS}">` +
      '<pr:Relationship Id="rId1" Type="t" Target="a.xml"/>' +
      "</pr:Relationships>";
    expect(
      splicePart(xml, {
        root: "Relationships",
        append: '<pr:Relationship Id="rId2" Type="t" Target="b.xml"/>',
      })
    ).toBe(
      `${PROLOG}<pr:Relationships xmlns:pr="${RELS_NS}">` +
        '<pr:Relationship Id="rId1" Type="t" Target="a.xml"/>' +
        '<pr:Relationship Id="rId2" Type="t" Target="b.xml"/>' +
        "</pr:Relationships>"
    );
  });

  it("opens a self-closing root to hold the new children", () => {
    expect(
      splicePart(`${PROLOG}<w15:people ${xmlnsDecl("w15")}/>\n`, {
        root: "people",
        append: '<w15:person w15:author="Ada"/>',
      })
    ).toBe(
      `${PROLOG}<w15:people ${xmlnsDecl("w15")}>` +
        '<w15:person w15:author="Ada"/></w15:people>\n'
    );
  });

  it("ignores a closing tag inside a comment", () => {
    const xml =
      `<Relationships xmlns="${RELS_NS}">` +
      "<!-- </Relationships> -->" +
      '<Relationship Id="rId1" Type="t" Target="a.xml"/>' +
      "</Relationships>";
    expect(
      splicePart(xml, { root: "Relationships", append: "<Relationship/>" })
    ).toBe(
      `<Relationships xmlns="${RELS_NS}">` +
        "<!-- </Relationships> -->" +
        '<Relationship Id="rId1" Type="t" Target="a.xml"/>' +
        "<Relationship/></Relationships>"
    );
  });

  it("reads the opening tag past a > inside an attribute value", () => {
    expect(
      splicePart('<Types note="a > b"><Default Extension="xml"/></Types>', {
        root: "Types",
        prepend: '<Default Extension="png"/>',
      })
    ).toBe(
      '<Types note="a > b"><Default Extension="png"/><Default Extension="xml"/></Types>'
    );
  });

  it("inserts at the spot CHILD_ORDER lays down for numbering", () => {
    const xml =
      "<w:numbering>" +
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"/></w:abstractNum>' +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      '<w:numIdMacAtCleanup w:val="1"/>' +
      "</w:numbering>";
    expect(
      splicePart(xml, {
        root: "numbering",
        insert: [
          { name: "num", xml: '<w:num w:numId="2"/>' },
          { name: "abstractNum", xml: '<w:abstractNum w:abstractNumId="1"/>' },
        ],
      })
    ).toBe(
      "<w:numbering>" +
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"/></w:abstractNum>' +
        '<w:abstractNum w:abstractNumId="1"/>' +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
        '<w:num w:numId="2"/>' +
        '<w:numIdMacAtCleanup w:val="1"/>' +
        "</w:numbering>"
    );
  });

  it("replaces every child and keeps the prolog", () => {
    const xml =
      `${PROLOG}\n<!-- kept --><w:comments ${xmlnsDecl("w")}>` +
      '<w:comment w:id="1"/><!-- gone --><w:comment w:id="2"/>' +
      "</w:comments>\n";
    expect(
      splicePart(xml, {
        root: "comments",
        replaceChildren: '<w:comment w:id="3"/>',
      })
    ).toBe(
      `${PROLOG}\n<!-- kept --><w:comments ${xmlnsDecl("w")}>` +
        '<w:comment w:id="3"/></w:comments>\n'
    );
  });

  it("throws malformed-xml naming the part when the root never closes", () => {
    const unclosed = `<w:numbering ${xmlnsDecl("w")}><w:num w:numId="1"/>`;
    expect(refusal(() => splicePart(unclosed, { root: "numbering" }))).toEqual({
      code: "malformed-xml",
      message: "the numbering part has no closing numbering tag",
    });
    expect(partRootProblem(unclosed, "numbering")).toBe(
      "the numbering part has no closing numbering tag"
    );
  });

  it("throws malformed-xml naming the part when the root is some other element", () => {
    const other = `${PROLOG}<w:notes ${xmlnsDecl("w")}><w:comment w:id="1"/></w:notes>`;
    expect(refusal(() => splicePart(other, { root: "comments" }))).toEqual({
      code: "malformed-xml",
      message: "the comments part has no comments root element",
    });
    expect(partRootProblem(other, "comments")).toBe(
      "the comments part has no comments root element"
    );
    expect(partRootProblem(other, "notes")).toBeNull();
  });
});

describe("ensureRootDeclarations", () => {
  const THREAD_MARKUP: RootDeclarations = {
    namespaces: { w14: NAMESPACES.w14, mc: NAMESPACES.mc },
    ignorable: ["w14"],
  };

  it("declares a namespace and an mc:Ignorable token once", () => {
    const bare = `<w:comments ${xmlnsDecl("w")}><w:comment/></w:comments>`;
    const declared = ensureRootDeclarations(bare, THREAD_MARKUP);
    expect(declared).toBe(
      `<w:comments ${xmlnsDecl("w")} ${xmlnsDecl("w14")} ${xmlnsDecl("mc")} mc:Ignorable="w14">` +
        "<w:comment/></w:comments>"
    );
    expect(ensureRootDeclarations(declared, THREAD_MARKUP)).toBe(declared);
  });

  it("adds a token to the mc:Ignorable the root already carries, keeping its quoting", () => {
    const ignoring =
      `<w:comments ${xmlnsDecl("w")} ${xmlnsDecl("mc")} mc:Ignorable='w15'>` +
      "<w:comment/></w:comments>";
    expect(ensureRootDeclarations(ignoring, THREAD_MARKUP)).toBe(
      `<w:comments ${xmlnsDecl("w")} ${xmlnsDecl("mc")} mc:Ignorable='w15 w14' ${xmlnsDecl("w14")}>` +
        "<w:comment/></w:comments>"
    );
  });

  it("declares on a root that closes on itself without opening it", () => {
    expect(
      ensureRootDeclarations(`<w:comments ${xmlnsDecl("w")}/>`, THREAD_MARKUP)
    ).toBe(
      `<w:comments ${xmlnsDecl("w")} ${xmlnsDecl("w14")} ${xmlnsDecl("mc")} mc:Ignorable="w14"/>`
    );
  });
});
