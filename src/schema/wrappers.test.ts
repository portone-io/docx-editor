// @vitest-environment jsdom
import type { Mark, MarkType, NodeType } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { docxSchema } from "./index";
import {
  innermostDepth,
  isWrapperType,
  sharedWrappers,
  WRAPPER_GROUP,
  wrapperMarks,
  wrapperOf,
  wrappersOf,
} from "./wrappers";

const WRAPPER_TYPES: readonly MarkType[] = Object.values(
  docxSchema.marks
).filter(isWrapperType);

const control = (attrs: { depth?: number; key?: number }): Mark =>
  docxSchema.marks.sdt.create({ sdtPrefix: "<w:sdt>", ...attrs });

const link = (attrs: { depth?: number; key?: number }): Mark =>
  docxSchema.marks.link.create({ linkPrefix: "<w:hyperlink>", ...attrs });

const textWearing = (...marks: Mark[]) => docxSchema.text("x", marks);

describe("the wrappers a node stands inside", () => {
  /**
   * `Mark.addToSet` puts a mark of the same type wherever the set already runs, so two controls
   * come back in the order they were added rather than the order the file wrote them. The depth is
   * what the file said, so it wins.
   */
  it("wrapperMarks orders by depth then declaration rank", () => {
    const inner = control({ depth: 1, key: 1 });
    const outer = control({ depth: 0, key: 0 });
    const added = outer.addToSet(inner.addToSet([]));
    expect(added.map((mark) => mark.attrs.depth)).toEqual([1, 0]);

    const node = textWearing(...added);
    expect(wrapperMarks(node).map((mark) => mark.attrs.key)).toEqual([0, 1]);

    // Two wrappers written at one depth fall back to the order the schema declares them in
    const tied = textWearing(link({ depth: 0 }), control({ depth: 0 }));
    expect(wrapperMarks(tied).map((mark) => mark.type.name)).toEqual([
      "sdt",
      "link",
    ]);
  });

  it("leaves out everything that is not a wrapper", () => {
    const node = textWearing(control({}), docxSchema.marks.run.create({}));
    expect(wrapperMarks(node).map((mark) => mark.type.name)).toEqual(["sdt"]);
  });

  it("answers with the outermost wrapper of a kind, and with all of them", () => {
    const node = textWearing(
      control({ depth: 0, key: 0 }),
      link({ depth: 1 }),
      control({ depth: 2, key: 1 })
    );
    expect(wrapperOf(node, "sdt")?.attrs.key).toBe(0);
    expect(wrapperOf(node, "link")?.attrs.depth).toBe(1);
    expect(wrappersOf(node, "sdt").map((mark) => mark.attrs.key)).toEqual([
      0, 1,
    ]);
    expect(wrapperOf(docxSchema.text("x"), "sdt")).toBeNull();
  });

  it("says what depth a wrapper laid inside these takes", () => {
    expect(innermostDepth([])).toBe(-1);
    expect(innermostDepth([control({ depth: 0 }), link({ depth: 3 })])).toBe(3);
    // A run mark is no wrapper, so it settles no depth
    expect(innermostDepth([docxSchema.marks.run.create({})])).toBe(-1);
  });

  it("shares the wrappers every node of a stretch stands inside alike", () => {
    const outer = control({ depth: 0, key: 0 });
    const inside = link({ depth: 1 });
    const other = link({ depth: 1, key: 1 });

    expect(
      sharedWrappers([
        textWearing(outer, inside),
        textWearing(outer, inside),
      ]).map((mark) => mark.attrs.depth)
    ).toEqual([0, 1]);
    // The second node stands in the same control but in a link of its own
    expect(
      sharedWrappers([textWearing(outer, inside), textWearing(outer, other)])
    ).toHaveLength(1);
    expect(
      sharedWrappers([textWearing(outer), docxSchema.text("y")])
    ).toHaveLength(0);
    expect(sharedWrappers([])).toHaveLength(0);
  });
});

describe("the schema's wrapper group", () => {
  const INLINE_NODES: readonly NodeType[] = Object.values(
    docxSchema.nodes
  ).filter((type) => type.isInline && type.name !== "text");

  it("holds the content control and the hyperlink", () => {
    expect(WRAPPER_TYPES.map((type) => type.name)).toEqual(["sdt", "link"]);
    expect(isWrapperType(docxSchema.marks.run)).toBe(false);
  });

  /**
   * A new wrapper arrives as one mark spec, so no node may name the kinds it takes one by one:
   * the whitelist names the group, and every inline node takes the whole of it.
   */
  it("every inline node spec allows the wrapper group and nothing else of it", () => {
    expect(INLINE_NODES.length).toBeGreaterThan(0);
    for (const type of INLINE_NODES) {
      expect([
        type.name,
        WRAPPER_TYPES.filter((wrapper) => type.allowsMarkType(wrapper)).map(
          (wrapper) => wrapper.name
        ),
      ]).toEqual([type.name, WRAPPER_TYPES.map((wrapper) => wrapper.name)]);

      const named = (type.spec.marks ?? "").split(" ");
      expect([type.name, named.includes(WRAPPER_GROUP)]).toEqual([
        type.name,
        true,
      ]);
      expect([
        type.name,
        WRAPPER_TYPES.filter((wrapper) => named.includes(wrapper.name)),
      ]).toEqual([type.name, []]);
    }
  });

  /** §17.5.2.17 has a `w:sdt` hold a `w:sdt`, which a mark excluding its own kind could not record */
  it("lets one content control stand inside another", () => {
    const both = control({ depth: 1, key: 1 }).addToSet([
      control({ depth: 0, key: 0 }),
    ]);
    expect(both).toHaveLength(2);
    // A link still replaces a link: one hyperlink inside another is not a shape OOXML writes
    expect(link({ key: 1 }).addToSet([link({ key: 0 })])).toHaveLength(1);
  });
});
