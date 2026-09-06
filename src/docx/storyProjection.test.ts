// @vitest-environment jsdom
import { zipSync } from "fflate";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import {
  fixtureNames,
  makeDocx,
  readFixture,
  TINY_PNG_DATA_URL,
} from "../__testing__/docx";
import { docxSchema } from "../schema";
import { importDocx } from "./importDocx";
import { serializeBlock } from "./serializeBlock";
import type { SessionStore } from "./session";
import {
  comparableStory,
  isModelledBlock,
  type Story,
} from "./storyProjection";

/** The same package with one block standing in for the whole body, so a block can be reopened where it came from */
function reopenIn(session: SessionStore, blockXml: string): Story {
  const parts: Record<string, Uint8Array> = {};
  for (const [path, bytes] of session.parts) parts[path] = bytes;
  parts[session.mainPartPath] = new TextEncoder().encode(
    session.documentPrefix + blockXml + session.documentSuffix
  );
  return importDocx(zipSync(parts));
}

const asIs = (node: PMNode): PMNode => node;

describe("the writer as a fixed point", () => {
  it.each(fixtureNames)(
    "writes every modelled block of %s the same the second time around",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      let modelled = 0;
      doc.forEach((block) => {
        if (!isModelledBlock(block)) return;
        modelled += 1;
        const once = serializeBlock(block, session);
        const reopened = reopenIn(session, once);
        expect(reopened.doc.child(0).type.name).toBe(block.type.name);
        expect(serializeBlock(reopened.doc.child(0), reopened.session)).toBe(
          once
        );
      });
      expect(modelled).toBeGreaterThan(0);
    }
  );
});

const CELL = (width: string) =>
  `<w:tc><w:tcPr>${width}</w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>`;

const TABLE = (tblW: string, tcW: string, grid: string) =>
  `<w:tbl><w:tblPr>${tblW}</w:tblPr>${grid}<w:tr>${CELL(tcW)}</w:tr></w:tbl>`;

const GRID = '<w:tblGrid><w:gridCol w:w="6500"/></w:tblGrid>';

function storyOf(body: string): readonly string[] {
  const projected = comparableStory(importDocx(makeDocx(body)), asIs);
  if (projected === null) throw new Error("the body could not be written back");
  return projected;
}

/**
 * What the comparison forgives, one case each.
 *
 * Writing both sides out through the same writer takes away every difference the writer itself
 * levels, and these are those differences. The list is the documented limit of the comparison, so
 * a lexical difference outside it is not forgiven and adding a case here is a decision to review.
 */
describe("accepted losses", () => {
  it("reads a Word-ordered tcW and tblW and a writer-ordered one as the same block", () => {
    expect(
      storyOf(
        TABLE(
          '<w:tblW w:type="dxa" w:w="6500"/>',
          '<w:tcW w:type="dxa" w:w="6500"/>',
          GRID
        )
      )
    ).toEqual(
      storyOf(
        TABLE(
          '<w:tblW w:w="6500" w:type="dxa"/>',
          '<w:tcW w:w="6500" w:type="dxa"/>',
          GRID
        )
      )
    );
  });

  it("reads a pct width written as 100% and one written as 5000 as the same block", () => {
    expect(
      storyOf(
        TABLE(
          '<w:tblW w:w="100%" w:type="pct"/>',
          '<w:tcW w:w="100%" w:type="pct"/>',
          GRID
        )
      )
    ).toEqual(
      storyOf(
        TABLE(
          '<w:tblW w:w="5000" w:type="pct"/>',
          '<w:tcW w:w="5000" w:type="pct"/>',
          GRID
        )
      )
    );
  });

  it("does not tell a table whose tblGridChange was dropped from the original", () => {
    const width = '<w:tblW w:w="6500" w:type="dxa"/>';
    const cell = '<w:tcW w:w="6500" w:type="dxa"/>';
    const revised = storyOf(
      TABLE(
        width,
        cell,
        '<w:tblGrid><w:gridCol w:w="6500"/>' +
          '<w:tblGridChange w:id="0"><w:tblGrid><w:gridCol w:w="4000"/>' +
          "</w:tblGrid></w:tblGridChange></w:tblGrid>"
      )
    );
    expect(revised.join("")).not.toContain("tblGridChange");
    expect(revised).toEqual(storyOf(TABLE(width, cell, GRID)));
  });

  it("tells two tables whose cells say different things apart", () => {
    const width = '<w:tblW w:w="6500" w:type="dxa"/>';
    const cell = '<w:tcW w:w="6500" w:type="dxa"/>';
    expect(storyOf(TABLE(width, cell, GRID))).not.toEqual(
      storyOf(TABLE(width, cell, GRID).replace("<w:t>a</w:t>", "<w:t>b</w:t>"))
    );
  });
});

describe("a story the writer cannot put out", () => {
  it("answers with nothing to compare rather than throwing", () => {
    const { doc, session } = importDocx(
      makeDocx("<w:p><w:r><w:t>a</w:t></w:r></w:p>")
    );
    // An image inserted during editing points at a relationship the projection has none of
    const inserted = docxSchema.nodes.image.create({
      src: TINY_PNG_DATA_URL,
      extent: { cx: 1905000, cy: 952500 },
    });
    const withImage = doc.copy(
      Fragment.from(doc.child(0).copy(Fragment.from(inserted)))
    );
    expect(comparableStory({ doc: withImage, session }, asIs)).toBeNull();
  });
});
