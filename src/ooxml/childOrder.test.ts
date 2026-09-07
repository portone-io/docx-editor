// @vitest-environment jsdom
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CHILD_ORDER, childOrderOf } from "./childOrder";

const XSD_NS = "http://www.w3.org/2001/XMLSchema";
const schema = new DOMParser().parseFromString(
  readFileSync(
    join(
      dirname(fileURLToPath(import.meta.url)),
      "../../spec/schemas/transitional/wml.xsd"
    ),
    "utf8"
  ),
  "application/xml"
);

function declaration(tag: string, name: string): Element {
  const found = Array.from(schema.documentElement.children).find(
    (el) => el.localName === tag && el.getAttribute("name") === name
  );
  if (!found) throw new Error(`wml.xsd declares no ${tag} named ${name}`);
  return found;
}

function particleName(el: Element): string {
  const name = el.getAttribute("name") ?? el.getAttribute("ref");
  if (name === null) throw new Error("an element particle names nothing");
  return name.slice(name.indexOf(":") + 1);
}

/**
 * Expand the content model in schema document order. A repeating choice does not require this
 * order, but the registry deliberately uses it as its writer convention. Unsupported content
 * models must fail here instead of silently yielding an incomplete list.
 */
function elements(el: Element): Element[] {
  switch (el.localName) {
    case "element":
      return [el];
    case "group": {
      const ref = el.getAttribute("ref");
      return ref === null
        ? Array.from(el.children).flatMap(elements)
        : elements(declaration("group", ref));
    }
    case "extension":
      return [
        ...elements(declaration("complexType", el.getAttribute("base") ?? "")),
        ...Array.from(el.children).flatMap(elements),
      ];
    case "complexType":
    case "complexContent":
    case "sequence":
    case "choice":
      return Array.from(el.children).flatMap(elements);
    // These declarations describe attributes or documentation, never child elements.
    case "attribute":
    case "attributeGroup":
    case "anyAttribute":
    case "annotation":
      return [];
    default:
      throw new Error(`unsupported XSD content: ${el.localName}`);
  }
}

function typeElements(type: string): Element[] {
  return elements(declaration("complexType", type));
}

function declares(el: Element, name: string, type: string): boolean {
  return particleName(el) === name && el.getAttribute("type") === type;
}

describe("the child order registry", () => {
  it.each(Object.entries(CHILD_ORDER))(
    "%s names a declared element and follows its type's schema document order",
    (key, { type, order }) => {
      const path = key.split("/");
      const name = path[path.length - 1];
      const candidates =
        path.length === 1
          ? Array.from(schema.getElementsByTagNameNS(XSD_NS, "element"))
          : typeElements(CHILD_ORDER[path[0]].type);
      expect(
        candidates.some((el) => declares(el, name, type)),
        `${key}: ${type} is not its declared type`
      ).toBe(true);
      expect(order).toEqual(typeElements(type).map(particleName));
      expect(new Set(order).size).toBe(order.length);
    }
  );

  it.each(["any", "all"])(
    "refuses an unsupported xsd:%s inside a content model",
    (particle) => {
      const type = new DOMParser().parseFromString(
        `<xsd:complexType xmlns:xsd="${XSD_NS}"><xsd:sequence><xsd:${particle}/></xsd:sequence></xsd:complexType>`,
        "application/xml"
      ).documentElement;
      expect(() => elements(type)).toThrow(
        `unsupported XSD content: ${particle}`
      );
    }
  );

  it("resolves pPr/rPr ahead of rPr", () => {
    expect(childOrderOf("rPr", "pPr")).toBe(CHILD_ORDER["pPr/rPr"].order);
    expect(childOrderOf("rPr", "pPr").slice(0, 2)).toEqual(["ins", "del"]);
    expect(childOrderOf("rPr")).toBe(CHILD_ORDER.rPr.order);
    expect(childOrderOf("rPr", "lvl")).toBe(CHILD_ORDER.rPr.order);
  });

  it("throws for a parent that has no entry", () => {
    expect(() => childOrderOf("tblGrid")).toThrow(/no child order/);
    expect(() => childOrderOf("tblGrid", "tbl")).toThrow(/no child order/);
  });
});
