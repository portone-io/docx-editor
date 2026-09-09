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
import { NO_EXPORT_REFS } from "./exportRefs";
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
        const once = serializeBlock(block, { ...NO_EXPORT_REFS, session });
        const reopened = reopenIn(session, once);
        expect(reopened.doc.child(0).type.name).toBe(block.type.name);
        expect(
          serializeBlock(reopened.doc.child(0), {
            ...NO_EXPORT_REFS,
            session: reopened.session,
          })
        ).toBe(once);
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
const WIDTH = '<w:tblW w:w="6500" w:type="dxa"/>';
const CELL_WIDTH = '<w:tcW w:w="6500" w:type="dxa"/>';

function storyOf(body: string): readonly string[] {
  const projected = comparableStory(importDocx(makeDocx(body)), asIs);
  if (projected === null) throw new Error("the body could not be written back");
  return projected;
}

/**
 * Two spellings of one thing, one case each.
 *
 * Writing both sides out through the same writer takes the wording out of the comparison, and
 * these are the wordings known to fall under it. The set is open, since it is the writer that
 * decides it, and it may safely be: a difference that says nothing is nothing to report.
 */
describe("two spellings of the same thing", () => {
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

  it("reads runs split by a producer and one run saying the same as the same paragraph", () => {
    expect(storyOf("<w:p><w:r><w:t>ab</w:t></w:r></w:p>")).toEqual(
      storyOf("<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>")
    );
  });

  it("reads a w:t with and without xml:space as the same paragraph", () => {
    expect(storyOf("<w:p><w:r><w:t>a</w:t></w:r></w:p>")).toEqual(
      storyOf('<w:p><w:r><w:t xml:space="preserve">a</w:t></w:r></w:p>')
    );
  });

  it("reads a literal tab and a w:tab as the same paragraph", () => {
    expect(
      storyOf('<w:p><w:r><w:t xml:space="preserve">a\tb</w:t></w:r></w:p>')
    ).toEqual(
      storyOf("<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r></w:p>")
    );
  });
});

/**
 * Content this editor does not keep, one case each.
 *
 * These are not two spellings of one thing. The writer builds them afresh and what it read is
 * gone, so the comparison cannot report a difference it can no longer see. That makes this list
 * the one to watch: it stands for what the editor loses when it rebuilds a block, it shrinks as
 * the writer learns to carry more, and a case joining it is a preservation defect rather than a
 * comparison detail. `site/content/docs/core/verifying-a-commenters-file.mdx` names the same losses
 * to a reader.
 */
describe("content this editor does not keep", () => {
  it("reads a cell property the writer rebuilds from the model as the model says it", () => {
    // A vMerge on a cell nothing continues, and a gridSpan of one, say nothing the model records
    expect(storyOf(TABLE(WIDTH, CELL_WIDTH, GRID))).toEqual(
      storyOf(
        TABLE(
          WIDTH,
          CELL_WIDTH + '<w:gridSpan w:val="1"/><w:vMerge w:val="restart"/>',
          GRID
        )
      )
    );
  });

  it("does not tell a grid apart, which is built from the column widths alone", () => {
    expect(storyOf(TABLE(WIDTH, CELL_WIDTH, GRID))).toEqual(
      storyOf(
        TABLE(
          WIDTH,
          CELL_WIDTH,
          '<w:tblGrid><w:gridCol w:w="6500"><!-- unseen --></w:gridCol></w:tblGrid>'
        )
      )
    );
  });
});

/**
 * The other half of the rule. What the writer carries through is compared, so a rebuilt block is
 * not a place to put bytes the comparison cannot see.
 */
describe("differences the comparison sees", () => {
  it("tells two tables whose cells say different things apart", () => {
    expect(storyOf(TABLE(WIDTH, CELL_WIDTH, GRID))).not.toEqual(
      storyOf(
        TABLE(WIDTH, CELL_WIDTH, GRID).replace("<w:t>a</w:t>", "<w:t>b</w:t>")
      )
    );
  });

  it("tells a cell whose properties gained an XML comment apart", () => {
    expect(storyOf(TABLE(WIDTH, CELL_WIDTH, GRID))).not.toEqual(
      storyOf(TABLE(WIDTH, CELL_WIDTH + "<!-- smuggled -->", GRID))
    );
  });

  it("tells a cell whose properties gained stray text apart", () => {
    expect(storyOf(TABLE(WIDTH, CELL_WIDTH, GRID))).not.toEqual(
      storyOf(TABLE(WIDTH, CELL_WIDTH + "smuggled", GRID))
    );
  });

  it("tells a width element that gained a comment inside it apart", () => {
    // The model decides this element's attributes, so it is written afresh; what a producer put
    // inside it is not the model's and comes along
    expect(storyOf(TABLE(WIDTH, CELL_WIDTH, GRID))).not.toEqual(
      storyOf(
        TABLE(
          WIDTH,
          '<w:tcW w:w="6500" w:type="dxa"><!-- smuggled --></w:tcW>',
          GRID
        )
      )
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
