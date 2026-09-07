// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { Fragment } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  exportErrorCode,
  makeDocx,
  TINY_PNG,
  TINY_PNG_DATA_URL,
} from "../__testing__/docx";
import { encodeUtf8 } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { exportThroughPlanners } from "./exportDocx";
import { importDocx } from "./importDocx";
import { CONTENT_TYPES_PATH, contentTypeWriter } from "./packageParts";
import {
  assertPartsParse,
  type PartPlanContext,
  type PartPlanner,
  runPartPlanners,
} from "./partPlan";
import { relationshipWriter } from "./relationships";

const opened = () =>
  importDocx(
    makeDocx('<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>')
  );

/** A context over a package whose content types part declares nothing yet */
function emptyContext(): PartPlanContext {
  const types =
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>';
  return {
    relationships: relationshipWriter([]),
    contentTypes: contentTypeWriter(
      new Map([[CONTENT_TYPES_PATH, encodeUtf8(types, false)]])
    ),
  };
}

/** A planner that writes these parts whatever the document holds, and declares each as it goes */
function writing(
  name: string,
  parts: Record<string, string>,
  contentType = `application/${name}+xml`
): PartPlanner {
  return {
    name,
    plan: (_doc, _session, context) => {
      for (const path of Object.keys(parts)) {
        context.contentTypes.addOverride(path, contentType);
      }
      return new Map(
        Object.entries(parts).map(([path, xml]) => [
          path,
          encodeUtf8(xml, false),
        ])
      );
    },
  };
}

const silent: PartPlanner = { name: "silent", plan: () => null };

describe("runPartPlanners", () => {
  it("folds every planner's parts into one replacement map", () => {
    const { doc, session } = opened();
    const context = emptyContext();
    const folded = runPartPlanners(
      [
        writing("first", { "word/first.xml": "<first/>" }),
        silent,
        writing("second", {
          "word/second.xml": "<second/>",
          "word/third.xml": "<third/>",
        }),
      ],
      doc,
      session,
      context
    );
    expect(
      Array.from(folded, ([path, bytes]) => [
        path,
        new TextDecoder().decode(bytes),
      ])
    ).toEqual([
      ["word/first.xml", "<first/>"],
      ["word/second.xml", "<second/>"],
      ["word/third.xml", "<third/>"],
    ]);
    // Every planner declared through the one writer, which answers for all of them at once
    const declared = context.contentTypes.part();
    expect(declared).not.toBeNull();
    const types = new TextDecoder().decode(declared ?? new Uint8Array());
    expect(types).toContain('PartName="/word/first.xml"');
    expect(types).toContain('PartName="/word/third.xml"');
  });

  it("refuses a planner that writes a part an earlier planner wrote", () => {
    const { doc, session } = opened();
    expect(() =>
      runPartPlanners(
        [
          writing("first", { "word/shared.xml": "<first/>" }),
          writing("second", { "word/shared.xml": "<second/>" }),
        ],
        doc,
        session,
        emptyContext()
      )
    ).toThrow(/second planner wrote word\/shared\.xml/);
  });

  it("refuses a planner that writes a part the context's writers own", () => {
    const { doc, session } = opened();
    for (const path of [
      CONTENT_TYPES_PATH,
      "word/_rels/document.xml.rels",
      session.mainPartPath,
    ]) {
      expect(() =>
        runPartPlanners(
          [writing("rogue", { [path]: "<Types/>" })],
          doc,
          session,
          emptyContext()
        )
      ).toThrow(/rogue planner wrote/);
    }
  });
});

describe("the export over its planners", () => {
  it("refuses with malformed-xml naming the part a planner wrote that does not read back", () => {
    const { doc, session } = opened();
    const rogue: PartPlanner = {
      name: "rogue",
      plan: () =>
        new Map([["word/rogue.xml", encodeUtf8("<rogue><entry/>", false)]]),
    };
    expect(
      exportErrorCode(() => exportThroughPlanners([rogue], doc, session))
    ).toBe("malformed-xml");
    expect(() => exportThroughPlanners([rogue], doc, session)).toThrow(
      "word/rogue.xml as written could not be parsed"
    );
  });

  it("refuses a planner that overwrites media already written by the export", () => {
    const original = unzipSync(makeDocx("<w:p/>"));
    original[CONTENT_TYPES_PATH] = new TextEncoder().encode(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
    );
    const opened = importDocx(zipSync(original));
    const image = docxSchema.nodes.image.create({
      src: TINY_PNG_DATA_URL,
      extent: { cx: 9525, cy: 9525 },
    });
    const doc = opened.doc.copy(
      Fragment.from(opened.doc.child(0).copy(Fragment.from(image)))
    );
    const written = unzipSync(exportThroughPlanners([], doc, opened.session));
    const path = Object.keys(written).find((path) =>
      path.startsWith("word/media/")
    );
    if (!path) throw new Error("expected exported image part");
    expect(() =>
      exportThroughPlanners(
        [writing("rogue", { [path]: "<not-an-image/>" })],
        doc,
        opened.session
      )
    ).toThrow(/rogue planner wrote/);
  });

  it("reads a rewritten XML part identified by content type rather than its file extension", () => {
    const { doc, session } = opened();
    session.parts.set(
      CONTENT_TYPES_PATH,
      encodeUtf8(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        false
      )
    );
    expect(() =>
      exportThroughPlanners(
        [writing("renamed", { "word/custom.data": "<broken>" })],
        doc,
        session
      )
    ).toThrow("word/custom.data as written could not be parsed");
  });

  it("repacks a part a planner wrote that reads back", () => {
    const { doc, session } = opened();
    const fine: PartPlanner = {
      name: "fine",
      plan: () => new Map([["word/fine.xml", encodeUtf8("<fine/>", false)]]),
    };
    const parts = unzipSync(exportThroughPlanners([fine], doc, session));
    expect(new TextDecoder().decode(parts["word/fine.xml"])).toBe("<fine/>");
  });
});

describe("assertPartsParse", () => {
  it("refuses to repack when a rewritten part does not parse", () => {
    const parts = new Map<string, Uint8Array>([
      ["word/media/image1.png", TINY_PNG],
      ["word/comments.xml", encodeUtf8("<w:comments><w:comment/>", false)],
    ]);
    expect(exportErrorCode(() => assertPartsParse(parts))).toBe(
      "malformed-xml"
    );
    expect(() => assertPartsParse(parts)).toThrow(
      "word/comments.xml as written could not be parsed"
    );
  });

  it("reads a relationships part back as XML too, and passes over a media part", () => {
    expect(() =>
      assertPartsParse(
        new Map([
          [
            "word/_rels/document.xml.rels",
            encodeUtf8("<Relationships>", false),
          ],
        ])
      )
    ).toThrow("word/_rels/document.xml.rels as written could not be parsed");
    expect(() =>
      assertPartsParse(
        new Map([
          ["word/media/image1.png", encodeUtf8("<not xml", false)],
          [
            "word/numbering.xml",
            encodeUtf8('<w:numbering xmlns:w="x"/>', true),
          ],
        ])
      )
    ).not.toThrow();
  });
});
