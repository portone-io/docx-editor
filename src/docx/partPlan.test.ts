// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { exportErrorCode, makeDocx, TINY_PNG } from "../__testing__/docx";
import { encodeUtf8 } from "../ooxml/xml";
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
    for (const path of [CONTENT_TYPES_PATH, "word/_rels/document.xml.rels"]) {
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
