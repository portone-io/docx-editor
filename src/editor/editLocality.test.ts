// @vitest-environment jsdom
/**
 * The one promise every editing move makes about the file it came from: the block the move was
 * aimed at is the only one rebuilt, and the original bytes stand untouched on either side of it.
 *
 * What each move does is checked where that move lives. What is asked here is only that promise,
 * of every move against every fixture.
 */

import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  bytesEqual,
  documentXmlOf,
  fixtureNames,
  readFixture,
  surroundings,
} from "../__testing__/docx";
import { runCommand, select } from "../__testing__/editing";
import { exportDocx } from "../docx/exportDocx";
import { importDocx } from "../docx/importDocx";
import type { LineSpacing, ParagraphAlign } from "../model/format";
import { setFontFamily, toggleBold } from "./commands/formattingCommands";
import {
  activeParagraphAlign,
  setParagraphAlign,
} from "./commands/paragraphCommands";
import { setLineSpacing } from "./commands/spacingCommands";
import { createEditorState } from "./createEditor";
import { docxKeymap } from "./plugins/keymap";

const DOUBLE: LineSpacing = { rule: "auto", lines: 2 };

/** The first paragraph that has text, and the range its first piece of text covers */
function firstTextSpot(doc: PMNode): {
  index: number;
  from: number;
  to: number;
} {
  let spot: { index: number; from: number; to: number } | null = null;
  doc.forEach((block, blockOffset, index) => {
    if (spot !== null || block.type.name !== "paragraph") return;
    block.forEach((child, childOffset) => {
      if (spot !== null || !child.isText) return;
      const start = blockOffset + 1 + childOffset;
      spot = { index, from: start + 1, to: start + child.nodeSize };
    });
  });
  if (spot === null) throw new Error("no paragraph with text");
  return spot;
}

/** The first paragraph block, text or not, and a position inside it */
function firstParagraph(doc: PMNode): { index: number; pos: number } {
  let found: { index: number; pos: number } | null = null;
  doc.forEach((block, offset, index) => {
    if (found !== null || block.type.name !== "paragraph") return;
    found = { index, pos: offset + 1 };
  });
  if (found === null) throw new Error("no paragraph");
  return found;
}

/** A value different from the alignment in force, so that there is always an edit to make */
function otherAlign(state: EditorState): ParagraphAlign {
  const current = activeParagraphAlign(state);
  return current.kind === "shared" && current.align === "center"
    ? "right"
    : "center";
}

interface LocalEdit {
  /** Which body block the move is aimed at, and the document it leaves behind */
  edit: (doc: PMNode, state: EditorState) => { index: number; doc: PMNode };
  /** What the one rebuilt block has to show, for a move that leaves a mark to recognize it by */
  rebuilt?: (block: string) => void;
}

const LOCAL_EDITS: ReadonlyArray<[string, LocalEdit]> = [
  [
    "typing a character",
    {
      edit: (doc, state) => {
        const spot = firstTextSpot(doc);
        return {
          index: spot.index,
          doc: state.apply(state.tr.insertText("inserted", spot.from)).doc,
        };
      },
      rebuilt: (block) => expect(block).toContain("inserted"),
    },
  ],
  [
    "splitting the paragraph with Enter",
    {
      edit: (doc, state) => {
        const spot = firstTextSpot(doc);
        const split = runCommand(select(state, spot.from), docxKeymap.Enter);
        expect(split.doc.childCount).toBe(doc.childCount + 1);
        return { index: spot.index, doc: split.doc };
      },
      rebuilt: (block) => expect(block.match(/<w:p[ >]/g)).toHaveLength(2),
    },
  ],
  [
    "turning the text bold",
    {
      edit: (doc, state) => {
        const spot = firstTextSpot(doc);
        const bolded = runCommand(
          select(state, spot.from, spot.to),
          toggleBold
        );
        return { index: spot.index, doc: bolded.doc };
      },
    },
  ],
  [
    "changing the font",
    {
      edit: (doc, state) => {
        const spot = firstTextSpot(doc);
        const changed = runCommand(
          select(state, spot.from, spot.to),
          setFontFamily("Nanum Gothic")
        );
        return { index: spot.index, doc: changed.doc };
      },
    },
  ],
  [
    "changing the alignment",
    {
      edit: (doc, state) => {
        const first = firstParagraph(doc);
        const spot = select(state, first.pos);
        const aligned = runCommand(spot, setParagraphAlign(otherAlign(spot)));
        return { index: first.index, doc: aligned.doc };
      },
    },
  ],
  [
    "changing the line spacing",
    {
      edit: (doc, state) => {
        const first = firstParagraph(doc);
        const spaced = runCommand(
          select(state, first.pos),
          setLineSpacing(DOUBLE)
        );
        return { index: first.index, doc: spaced.doc };
      },
    },
  ],
];

describe.each(LOCAL_EDITS)("%s", (_move, move) => {
  it.each(fixtureNames)(
    "%s: rebuilds only the block it was aimed at",
    (name) => {
      const { doc, session } = importDocx(readFixture(name));
      const state = createEditorState(doc, { styles: session.styles });
      const edited = move.edit(doc, state);

      const documentXml = documentXmlOf(edited.doc, session);
      const { head, tail } = surroundings(session, edited.index);
      expect(documentXml.startsWith(head)).toBe(true);
      expect(documentXml.endsWith(tail)).toBe(true);
      move.rebuilt?.(
        documentXml.slice(head.length, documentXml.length - tail.length)
      );
    }
  );
});

describe("the parts outside the body", () => {
  it.each(fixtureNames)("%s: are left exactly as they were", (name) => {
    const bytes = readFixture(name);
    const { doc, session } = importDocx(bytes);
    const state = createEditorState(doc, { styles: session.styles });
    const edited = state.apply(
      state.tr.insertText("inserted", firstTextSpot(doc).from)
    );

    const original = unzipSync(bytes);
    const exported = unzipSync(exportDocx(edited.doc, session));
    for (const key of Object.keys(original)) {
      if (key === session.mainPartPath) continue;
      expect(bytesEqual(exported[key], original[key])).toBe(true);
    }
  });
});
