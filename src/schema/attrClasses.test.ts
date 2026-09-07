// @vitest-environment jsdom
/** Checks attr provenance against the schema and the direct reads of the two block writers. */
import { readFileSync } from "node:fs";
import type { MarkType, NodeType } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  type AttrClass,
  type AttrFacts,
  attrsOfClass,
  MARK_ATTR_ROLES,
  NODE_ATTR_ROLES,
} from "./attrRoles";
import { docxSchema } from "./index";

type Classified = readonly [
  NodeType | MarkType,
  Readonly<Record<string, AttrFacts>>,
];

function withAttrs(
  types: Readonly<Record<string, NodeType | MarkType>>,
  table: Readonly<Record<string, Readonly<Record<string, AttrFacts>>>>
): readonly Classified[] {
  return Object.values(types)
    .filter((type) => Object.keys(type.spec.attrs ?? {}).length > 0)
    .map((type) => [type, table[type.name] ?? {}] as const);
}

const NODES = withAttrs(docxSchema.nodes, NODE_ATTR_ROLES);
const MARKS = withAttrs(docxSchema.marks, MARK_ATTR_ROLES);

const attrsOf = (type: NodeType | MarkType) =>
  Object.keys(type.spec.attrs ?? {}).sort();

const CLASSES: readonly AttrClass[] = [
  "preserved",
  "derived",
  "identity",
  "model",
];

/** Every classified attr as the type it sits on, its name, and what the table says about it */
function everyAttr(): readonly (readonly [string, string, AttrFacts])[] {
  return [...NODES, ...MARKS].flatMap(([type, facts]) =>
    Object.entries(facts).map(
      ([name, fact]) => [type.name, name, fact] as const
    )
  );
}

/**
 * The attrs a writer reads, read out of its source the way the list was first drawn up. A writer
 * that reaches for an attr through a destructuring is invisible here, so this is a floor on what
 * the writers touch rather than the whole of it.
 */
function attrsReadBy(file: string): readonly string[] {
  const source = readFileSync(new URL(file, import.meta.url), "utf8");
  const names = [...source.matchAll(/\battrs\.([A-Za-z_]\w*)/g)].map(
    (match) => match[1]
  );
  return [...new Set(names)].sort();
}

/**
 * What the two block writers were reading when this table was drawn. They stand here so that a
 * writer reaching for a new attr fails this file, where the table it has to be classified in is.
 */
const SERIALIZE_TABLE_READS: readonly string[] = [
  "colspan",
  "gridChange",
  "gridCols",
  "rowspan",
  "sdtPrefix",
  "tblAttrs",
  "tblPr",
  "tblPrEx",
  "tblW",
  "tcAttrs",
  "tcPr",
  "tcW",
  "trAttrs",
  "trPr",
];

const SERIALIZE_PARAGRAPH_READS: readonly string[] = [
  "alt",
  "brAttrs",
  "extent",
  "href",
  "id",
  "kind",
  "linkPrefix",
  "pAttrs",
  "pPr",
  "rAttrs",
  "referenceXml",
  "rPr",
  "sdtPrefix",
  "src",
  "tabAttrs",
  "xml",
];

/** A name shaped like one of these carries markup, whichever node or mark declared it */
const PRESERVED_SUFFIXES = ["Pr", "Attrs", "Xml", "Prefix", "xml"];

describe("the classification covers the schema", () => {
  it.each(NODES)(
    "classifies every attr of $name exactly once",
    (type, facts) => {
      expect(Object.keys(facts).sort()).toEqual(attrsOf(type));
    }
  );

  it.each(MARKS)(
    "classifies every attr of the $name mark exactly once",
    (type, facts) => {
      expect(Object.keys(facts).sort()).toEqual(attrsOf(type));
    }
  );

  it("names no attr the schema does not have", () => {
    const declared = [...NODES, ...MARKS].flatMap(([type, facts]) =>
      Object.keys(facts).map((name) => `${type.name}.${name}`)
    );
    const inSchema = [...NODES, ...MARKS].flatMap(([type]) =>
      attrsOf(type).map((name) => `${type.name}.${name}`)
    );

    expect(declared.sort()).toEqual(inSchema.sort());
  });

  it("puts each attr in one of the four classes and no other", () => {
    const gathered = [...NODES, ...MARKS].flatMap(([type]) =>
      CLASSES.flatMap((cls) =>
        attrsOfClass(type, cls).map((name) => `${type.name}.${name}`)
      )
    );

    expect(new Set(gathered).size).toBe(gathered.length);
    expect(gathered.length).toBe(everyAttr().length);
  });
});

describe("the boundary the classification draws", () => {
  it.each([
    ["serializeTable", "../docx/serializeTable.ts", SERIALIZE_TABLE_READS],
    [
      "serializeParagraph",
      "../docx/serializeParagraph.ts",
      SERIALIZE_PARAGRAPH_READS,
    ],
  ])("%s reads the attrs this table was drawn against", (_, file, reads) => {
    expect(attrsReadBy(file)).toEqual([...reads].sort());
  });

  it("hands no derived attr to the serializers", () => {
    const derived = new Set(
      everyAttr()
        .filter(([, , fact]) => fact.class === "derived")
        .map(([, name]) => name)
    );

    expect(
      [...SERIALIZE_TABLE_READS, ...SERIALIZE_PARAGRAPH_READS].filter((name) =>
        derived.has(name)
      )
    ).toEqual([]);
  });

  it("classifies every attr the serializers read", () => {
    const classified = new Set(everyAttr().map(([, name]) => name));

    expect(
      [...SERIALIZE_TABLE_READS, ...SERIALIZE_PARAGRAPH_READS].filter(
        (name) => !classified.has(name)
      )
    ).toEqual([]);
  });

  it("calls every attr named after markup preserved", () => {
    const misfiled = everyAttr()
      .filter(([, name]) =>
        PRESERVED_SUFFIXES.some((suffix) => name.endsWith(suffix))
      )
      .filter(([, , fact]) => fact.class !== "preserved")
      .map(([type, name]) => `${type}.${name}`);

    expect(misfiled).toEqual([]);
  });
});

describe("reading a class", () => {
  it("answers what the table says", () => {
    expect(attrsOfClass(docxSchema.nodes.paragraph, "preserved")).toEqual([
      "pAttrs",
      "pPr",
    ]);
    expect(attrsOfClass(docxSchema.nodes.paragraph, "identity")).toEqual([
      "srcId",
    ]);
    expect(attrsOfClass(docxSchema.nodes.paragraph, "model")).toEqual([]);
  });

  it("reads the derived attrs of marks and nodes separately", () => {
    expect(attrsOfClass(docxSchema.marks.sdt, "derived")).toEqual([
      "contentsLocked",
      "deletionLocked",
    ]);
    expect(attrsOfClass(docxSchema.nodes.tableCell, "derived")).toEqual([
      "colwidth",
      "format",
      "sdtContentsLocked",
      "sdtDeletionLocked",
    ]);
  });
});
