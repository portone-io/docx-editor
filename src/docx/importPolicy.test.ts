// @vitest-environment jsdom
/**
 * The preservation table against the content models it was written from.
 *
 * The table is the whole vocabulary the readers know, so what it must not do is fall silent: an
 * element `wml.xsd` admits at a level and the table does not name would be preserved by the
 * level's default without anybody having decided that. Each level is expanded out of its own
 * complex type here, groups and all, and compared with the sub-table both ways round.
 */
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { M_NS, W_NS } from "../ooxml/names";
import { parseXml } from "../ooxml/xml";
import {
  type ContentLevel,
  clarkName,
  policyFor,
  runContentText,
  WML_POLICY,
} from "./importPolicy";

const specDir = join(
  dirname(fileURLToPath(import.meta.url)),
  "../../spec/schemas/transitional"
);

function schemaDocument(file: string): Document {
  return new DOMParser().parseFromString(
    readFileSync(join(specDir, file), "utf8"),
    "application/xml"
  );
}

const wml = schemaDocument("wml.xsd");
const math = schemaDocument("shared-math.xsd");

/** Every schema this test follows a reference into, by the namespace it declares */
const SCHEMAS: ReadonlyMap<string, Document> = new Map([
  [W_NS, wml],
  [M_NS, math],
]);

function declaration(document: Document, tag: string, name: string): Element {
  const found = Array.from(document.documentElement.children).find(
    (el) => el.localName === tag && el.getAttribute("name") === name
  );
  if (!found) throw new Error(`the schema declares no ${tag} named ${name}`);
  return found;
}

/** The namespace the schema declares its own elements in */
function targetNamespaceOf(document: Document): string {
  const declared = document.documentElement.getAttribute("targetNamespace");
  if (declared === null)
    throw new Error("the schema names no target namespace");
  return declared;
}

/** The namespace a prefix stands for on the schema's own root element */
function namespaceOf(document: Document, prefix: string): string {
  const declared = document.documentElement.getAttribute(`xmlns:${prefix}`);
  if (declared === null) {
    throw new Error(`the schema binds no prefix ${prefix}`);
  }
  return declared;
}

/**
 * The Clark name one element particle names.
 *
 * A particle carrying `name` declares an element of its own schema's namespace; one carrying
 * `ref` points at an element another schema declares, which is how `m:oMath` reaches a level.
 */
function particleClarkName(document: Document, el: Element): string {
  const name = el.getAttribute("name");
  if (name !== null) return `{${targetNamespaceOf(document)}}${name}`;
  const ref = el.getAttribute("ref");
  if (ref === null) throw new Error("an element particle names nothing");
  const colon = ref.indexOf(":");
  if (colon === -1) return `{${targetNamespaceOf(document)}}${ref}`;
  const namespace = namespaceOf(document, ref.slice(0, colon));
  const referenced = SCHEMAS.get(namespace);
  if (referenced === undefined) {
    throw new Error(`no committed schema declares ${ref}`);
  }
  // Following the reference is what proves the element exists rather than that a prefix parses
  declaration(referenced, "element", ref.slice(colon + 1));
  return `{${namespace}}${ref.slice(colon + 1)}`;
}

/** The content model expanded, with every group reference followed */
function elements(el: Element): Element[] {
  switch (el.localName) {
    case "element":
      return [el];
    case "group": {
      const ref = el.getAttribute("ref");
      return ref === null
        ? Array.from(el.children).flatMap(elements)
        : elements(declaration(wml, "group", ref));
    }
    case "extension":
      return [
        ...elements(
          declaration(wml, "complexType", el.getAttribute("base") ?? "")
        ),
        ...Array.from(el.children).flatMap(elements),
      ];
    case "complexType":
    case "complexContent":
    case "sequence":
    case "choice":
      return Array.from(el.children).flatMap(elements);
    case "attribute":
    case "attributeGroup":
    case "anyAttribute":
    case "annotation":
      return [];
    default:
      throw new Error(`unsupported XSD content: ${el.localName}`);
  }
}

/** The complex type each level is the content model of */
const LEVEL_TYPES: Readonly<Record<ContentLevel, string>> = {
  r: "CT_R",
  p: "CT_P",
  // The content of an inline control, which is `EG_PContent`, as `CT_Hyperlink` is
  wrapper: "CT_SdtContentRun",
  body: "CT_Body",
  tc: "CT_Tc",
  tbl: "CT_Tbl",
  tr: "CT_Row",
};

const LEVELS = Object.keys(LEVEL_TYPES).filter(
  (level): level is ContentLevel => level in LEVEL_TYPES
);

function admittedAt(level: ContentLevel): string[] {
  const names = elements(declaration(wml, "complexType", LEVEL_TYPES[level]))
    .map((el) => particleClarkName(wml, el))
    .sort();
  return Array.from(new Set(names));
}

/**
 * What each level admits and deliberately has no rule for, which is only ever the two levels with
 * no node to hold a stranger. Named rather than counted, so an element sliding into the demoting
 * default fails here.
 */
const DEMOTED: Readonly<Record<ContentLevel, readonly string[]>> = {
  r: [],
  p: [],
  wrapper: [],
  body: [],
  tc: [],
  tbl: [
    "customXml",
    "del",
    "ins",
    "moveFrom",
    "moveTo",
    "oMath",
    "oMathPara",
    "sdt",
  ],
  tr: ["customXml", "del", "ins", "moveFrom", "moveTo", "oMath", "oMathPara"],
};

/** An element of the wordprocessing namespace, built the way a reader meets one */
function element(localName: string, namespace = W_NS): Element {
  return parseXml(
    `<x xmlns:w="${W_NS}" xmlns:m="${M_NS}"><${namespace === W_NS ? "w" : "m"}:${localName}/></x>`
  ).documentElement.children[0];
}

function tierOf(localName: string, level: ContentLevel): string {
  return policyFor(element(localName), level).tier;
}

describe("the preservation table against wml.xsd", () => {
  it("reads the committed schemas", () => {
    expect(wml.documentElement.localName).toBe("schema");
    expect(math.documentElement.localName).toBe("schema");
  });

  it.each(LEVELS)(
    "%s: every element the schema admits at this level has a rule in its sub-table",
    (level) => {
      const missing = admittedAt(level)
        .filter((name) => !WML_POLICY[level].has(name))
        .map((name) => name.slice(name.indexOf("}") + 1))
        .sort();

      expect(missing).toEqual(Array.from(DEMOTED[level]).sort());
    }
  );

  it.each(LEVELS)(
    "%s: no rule in this sub-table names an element the schema does not admit here",
    (level) => {
      const admitted = new Set(admittedAt(level));
      const strangers = Array.from(WML_POLICY[level].keys())
        .filter((name) => !admitted.has(name))
        .sort();

      expect(strangers).toEqual([]);
    }
  );

  it("reaches the math elements through the schema that declares them", () => {
    expect(admittedAt("p")).toContain(`{${M_NS}}oMath`);
    expect(WML_POLICY.p.get(`{${M_NS}}oMath`)?.tier).toBe("inline");
    expect(policyFor(element("oMath", M_NS), "p").tier).toBe("inline");
  });

  it.each(["r", "p", "wrapper", "body", "tc"] as const)(
    "%s: an unknown element falls to the narrowest preservation tier of the level",
    (level) => {
      const narrowest = { r: "runContent", p: "inline", wrapper: "inline" };
      expect(tierOf("notInWml", level)).toBe(
        level === "body" || level === "tc" ? "block" : narrowest[level]
      );
    }
  );

  it.each(["tbl", "tr"] as const)(
    "%s: an unknown element demotes the table",
    (level) => {
      expect(tierOf("notInWml", level)).toBe("demote");
      expect(tierOf("customXml", level)).toBe("demote");
    }
  );

  it("sdt is a model at p, a block at body and tc, and a demotion at tbl", () => {
    expect(tierOf("sdt", "p")).toBe("model");
    expect(tierOf("sdt", "body")).toBe("block");
    expect(tierOf("sdt", "tc")).toBe("block");
    expect(tierOf("sdt", "tr")).toBe("model");
    expect(tierOf("sdt", "tbl")).toBe("demote");
  });

  it("names an element by its namespace rather than by its prefix alone", () => {
    expect(clarkName(element("sdt"))).toBe(`{${W_NS}}sdt`);
    expect(clarkName(element("oMath", M_NS))).toBe(`{${M_NS}}oMath`);
    // A `sdt` of another namespace is not the control the paragraph reader unwraps
    expect(policyFor(element("sdt", M_NS), "p").tier).toBe("inline");
  });

  it("refuses an unsupported xsd particle instead of yielding a short list", () => {
    const type = new DOMParser().parseFromString(
      '<xsd:complexType xmlns:xsd="http://www.w3.org/2001/XMLSchema">' +
        "<xsd:sequence><xsd:any/></xsd:sequence></xsd:complexType>",
      "application/xml"
    ).documentElement;

    expect(() => elements(type)).toThrow("unsupported XSD content: any");
  });
});

describe("the text a run child puts on screen", () => {
  function textOf(xml: string): string | null {
    return runContentText(
      parseXml(`<x xmlns:w="${W_NS}">${xml}</x>`).documentElement.children[0]
    );
  }

  it("reads a text element as what it holds", () => {
    expect(textOf('<w:t xml:space="preserve"> a </w:t>')).toBe(" a ");
    expect(textOf("<w:t></w:t>")).toBe("");
  });

  it("reads the characters the editor models as itself", () => {
    expect(textOf("<w:tab/>")).toBe("\t");
    expect(textOf("<w:br/>")).toBe("\n");
  });

  it("reads a carriage return as a line break and a no-break hyphen as its character", () => {
    expect(textOf("<w:cr/>")).toBe("\n");
    expect(textOf("<w:noBreakHyphen/>")).toBe("‑");
  });

  it("reads nothing for what puts nothing on screen", () => {
    expect(textOf("<w:softHyphen/>")).toBeNull();
    expect(textOf("<w:lastRenderedPageBreak/>")).toBeNull();
    expect(textOf('<w:sym w:font="Wingdings" w:char="F0E0"/>')).toBeNull();
    expect(textOf('<w:fldChar w:fldCharType="begin"/>')).toBeNull();
    expect(
      textOf('<w:drawing><wp:inline xmlns:wp="urn:x"/></w:drawing>')
    ).toBeNull();
  });

  it("reads nothing for an element that is no run child at all", () => {
    expect(textOf("<w:r><w:t>a</w:t></w:r>")).toBeNull();
  });
});
