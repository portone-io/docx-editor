// @vitest-environment jsdom
import type { MarkType, NodeType } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  attrRole,
  displayAttrsOf,
  MARK_ATTR_ROLES,
  NODE_ATTR_ROLES,
} from "./attrRoles";
import { docxSchema } from "./index";

/** Every type of the schema that carries attrs, beside the table that declares their roles */
const NODES: readonly [NodeType, Readonly<Record<string, string>>][] =
  Object.values(docxSchema.nodes)
    .filter((type) => Object.keys(type.spec.attrs ?? {}).length > 0)
    .map((type) => [type, NODE_ATTR_ROLES[type.name] ?? {}]);

const MARKS: readonly [MarkType, Readonly<Record<string, string>>][] =
  Object.values(docxSchema.marks)
    .filter((type) => Object.keys(type.spec.attrs ?? {}).length > 0)
    .map((type) => [type, MARK_ATTR_ROLES[type.name] ?? {}]);

const attrsOf = (type: NodeType | MarkType) =>
  Object.keys(type.spec.attrs ?? {}).sort();

describe("the roles cover the schema", () => {
  it.each(NODES)("every attr of $name has a role", (type, roles) => {
    expect(Object.keys(roles).sort()).toEqual(attrsOf(type));
  });

  it.each(MARKS)("every attr of the $name mark has a role", (type, roles) => {
    expect(Object.keys(roles).sort()).toEqual(attrsOf(type));
  });

  it("names no type the schema does not have", () => {
    expect(Object.keys(NODE_ATTR_ROLES).sort()).toEqual(
      NODES.map(([type]) => type.name).sort()
    );
    expect(Object.keys(MARK_ATTR_ROLES).sort()).toEqual(
      MARKS.map(([type]) => type.name).sort()
    );
  });
});

describe("reading a role", () => {
  it("answers what the table says", () => {
    expect(attrRole(docxSchema.nodes.paragraph, "pPr")).toBe("source");
    expect(attrRole(docxSchema.nodes.paragraph, "format")).toBe("display");
    expect(attrRole(docxSchema.nodes.paragraph, "srcId")).toBe("session");
    expect(attrRole(docxSchema.marks.run, "format")).toBe("display");
  });

  it("tells a mark from a node of the same name", () => {
    expect(attrRole(docxSchema.marks.sdt, "sdtPrefix")).toBe("source");
    expect(attrRole(docxSchema.nodes.tableCell, "sdtPrefix")).toBe("source");
  });

  it("reads an attr nobody declared as one the exporter writes from", () => {
    expect(attrRole(docxSchema.nodes.paragraph, "notAnAttr")).toBe("source");
  });

  it("gathers the derived attrs of a type", () => {
    expect([...displayAttrsOf(docxSchema.nodes.tableCell)].sort()).toEqual([
      "colwidth",
      "format",
    ]);
    expect(displayAttrsOf(docxSchema.nodes.rawInline)).toEqual([]);
  });
});
