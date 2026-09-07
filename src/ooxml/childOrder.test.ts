// @vitest-environment jsdom
/**
 * The registry against the schema it claims to mirror.
 *
 * `wml.xsd` is read as XML rather than validated with it, so this runs on a DOMParser alone and
 * needs no xmllint. What it expands is the part of XSD the WordprocessingML property types use:
 * a sequence, a choice, a `xsd:group ref`, and a `xsd:extension` of another complex type.
 */

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { CHILD_ORDER, childOrderOf } from "./childOrder";

const XSD_PATH = join(
  dirname(fileURLToPath(import.meta.url)),
  "../../spec/schemas/transitional/wml.xsd"
);

/**
 * One spot in a type's content model and the names that may stand in it.
 *
 * A sequence of elements gives one name per slot, which is a total order. A choice gives one slot
 * holding every alternative, since the schema says nothing about their order among themselves.
 */
interface Slot {
  names: string[];
  /** Whether the schema lets this slot come round more than once */
  repeats: boolean;
}

const schema = new DOMParser().parseFromString(
  readFileSync(XSD_PATH, "utf8"),
  "application/xml"
);

function declaration(tag: string, name: string): Element {
  const found = Array.from(schema.documentElement.children).find(
    (el) => el.localName === tag && el.getAttribute("name") === name
  );
  if (!found) throw new Error(`wml.xsd declares no ${tag} named ${name}`);
  return found;
}

function repeats(el: Element): boolean {
  const max = el.getAttribute("maxOccurs");
  return max !== null && max !== "1";
}

/** The local part of an element particle's name, whether it declares one or refers to one */
function particleName(el: Element): string {
  const named = el.getAttribute("name");
  if (named !== null) return named;
  const ref = el.getAttribute("ref");
  if (ref === null) throw new Error("an element particle names nothing");
  return ref.slice(ref.indexOf(":") + 1);
}

const PARTICLES = ["sequence", "choice", "group", "element"];

function particleSlots(el: Element, repeating: boolean): Slot[] {
  const repeated = repeating || repeats(el);
  switch (el.localName) {
    case "element":
      return [{ names: [particleName(el)], repeats: repeated }];
    case "sequence":
      return childSlots(el, repeated);
    case "choice":
      return [
        {
          names: childSlots(el, false).flatMap((slot) => slot.names),
          repeats: repeated,
        },
      ];
    case "group":
      return childSlots(
        declaration("group", el.getAttribute("ref") ?? ""),
        repeated
      );
    default:
      throw new Error(
        `wml.xsd holds a particle this test cannot expand: ${el.localName}`
      );
  }
}

function childSlots(el: Element, repeating: boolean): Slot[] {
  return Array.from(el.children)
    .filter((child) => PARTICLES.includes(child.localName))
    .flatMap((child) => particleSlots(child, repeating));
}

/** The content model of a complex type, with a base type it extends laid down first */
function typeSlots(name: string): Slot[] {
  const type = declaration("complexType", name);
  const content = Array.from(type.children).find(
    (child) => child.localName === "complexContent"
  );
  if (!content) return childSlots(type, false);
  const extension = Array.from(content.children).find(
    (child) => child.localName === "extension"
  );
  if (!extension) throw new Error(`${name} restricts rather than extends`);
  return [
    ...typeSlots(extension.getAttribute("base") ?? ""),
    ...childSlots(extension, false),
  ];
}

/**
 * The types whose content the schema orders from end to end: every slot holds one name, or holds
 * alternatives that cannot both appear. The entry has to be that order exactly.
 */
const SEQUENCE_TYPES: Readonly<Record<string, string>> = {
  pPr: "CT_PPr",
  tcPr: "CT_TcPr",
  tblPr: "CT_TblPr",
  tcBorders: "CT_TcBorders",
  tblBorders: "CT_TblBorders",
  pBdr: "CT_PBdr",
  tcMar: "CT_TcMar",
  tblCellMar: "CT_TblCellMar",
  sdt: "CT_SdtBlock",
  sdtPr: "CT_SdtPr",
  numPr: "CT_NumPr",
  lvl: "CT_Lvl",
  abstractNum: "CT_AbstractNum",
  num: "CT_Num",
  numbering: "CT_Numbering",
  tblStylePr: "CT_TblStylePr",
  style: "CT_Style",
  settings: "CT_Settings",
};

/**
 * The types holding a choice that may come round again, where the order among its alternatives is
 * this package's convention. The entry has to name the same set and to respect every boundary the
 * schema does fix: what stands before the choice, and what stands after it.
 */
const CHOICE_TYPES: Readonly<Record<string, string>> = {
  rPr: "CT_RPr",
  "pPr/rPr": "CT_ParaRPr",
  trPr: "CT_TrPr",
  sectPr: "CT_SectPr",
};

/** Where each name sits in the entry, so a slot's spread can be compared to the next slot's */
function positions(
  entry: readonly string[],
  names: readonly string[]
): number[] {
  return names.map((name) => {
    const at = entry.indexOf(name);
    if (at === -1) throw new Error(`the entry does not carry ${name}`);
    return at;
  });
}

describe("the child order registry", () => {
  it("checks every registered parent against the schema", () => {
    expect(Object.keys(CHILD_ORDER).toSorted()).toEqual(
      [...Object.keys(SEQUENCE_TYPES), ...Object.keys(CHOICE_TYPES)].toSorted()
    );
  });

  it("every sequence-typed entry equals the element order wml.xsd lays down", () => {
    for (const [parent, type] of Object.entries(SEQUENCE_TYPES)) {
      const slots = typeSlots(type);
      expect(
        slots.filter((slot) => slot.repeats && slot.names.length > 1),
        `${type} holds a choice that may repeat, so its order is a convention`
      ).toEqual([]);
      expect(CHILD_ORDER[parent], `${parent} does not follow ${type}`).toEqual(
        slots.flatMap((slot) => slot.names)
      );
    }
  });

  it("every choice-typed entry is a permutation of the schema's choice set", () => {
    for (const [parent, type] of Object.entries(CHOICE_TYPES)) {
      const slots = typeSlots(type);
      const entry = CHILD_ORDER[parent];
      expect(
        slots.some((slot) => slot.repeats && slot.names.length > 1),
        `${type} orders its children from end to end, so its entry can be compared exactly`
      ).toBe(true);
      expect(
        entry.toSorted(),
        `${parent} does not name what ${type} holds`
      ).toEqual(slots.flatMap((slot) => slot.names).toSorted());
      // Nothing crosses a boundary the schema does fix, whatever the order inside a choice is
      for (const [index, slot] of slots.slice(1).entries()) {
        const before = positions(entry, slots[index].names);
        const after = positions(entry, slot.names);
        expect(
          Math.max(...before),
          `${parent} writes ${slot.names.join("/")} before ${slots[index].names.join("/")}`
        ).toBeLessThan(Math.min(...after));
      }
    }
  });

  it("names each child of a parent exactly once", () => {
    for (const [parent, order] of Object.entries(CHILD_ORDER)) {
      expect(new Set(order).size, `${parent} names a child twice`).toBe(
        order.length
      );
    }
  });

  it("resolves pPr/rPr ahead of rPr", () => {
    expect(childOrderOf("rPr", "pPr")).toBe(CHILD_ORDER["pPr/rPr"]);
    expect(childOrderOf("rPr", "pPr").slice(0, 2)).toEqual(["ins", "del"]);
    expect(childOrderOf("rPr")).toBe(CHILD_ORDER.rPr);
    // A parent that registers nothing of its own falls back to the element's own entry
    expect(childOrderOf("rPr", "lvl")).toBe(CHILD_ORDER.rPr);
  });

  it("throws for a parent that has no entry", () => {
    expect(() => childOrderOf("tblGrid")).toThrow(/no child order/);
    expect(() => childOrderOf("tblGrid", "tbl")).toThrow(/no child order/);
  });
});
