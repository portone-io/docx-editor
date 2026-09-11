// @vitest-environment jsdom

import { unzipSync, zipSync } from "fflate";
import { TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { describe, it } from "vitest";
import {
  LETTER_SECT_PR,
  makeDocx,
  makeNotesDocx,
  makeStyledNumberedDocx,
} from "../../src/__testing__/docx";
import { exportDocx } from "../../src/docx/exportDocx";
import { importDocx } from "../../src/docx/importDocx";
import { sectionsOf } from "../../src/docx/sections";
import { noteProjection } from "../../src/editor/commands/noteQueries";
import { editorStateForSession } from "../../src/editor/createEditor";
import { noteNodeSpecs } from "../../src/editor/notes/noteSurface";
import { storyMarkup } from "../../src/editor/stories/storyMarkup";
import type { MeasuredBlock, MeasureTarget } from "../../src/page/blockKinds";
import type { DemandBand } from "../../src/page/demands";
import {
  FOOTNOTE_BAND,
  footnoteDemands,
} from "../../src/page/demands/footnoteDemands";
import { pageLayout, sectionPixels } from "../../src/page/pageLayout";
import { DEFAULT_FONT_FALLBACKS } from "../../src/styles/fontStack";

type Timings = {
  import: number;
  state: number;
  edit: number;
  export: number;
};

const sizes = [500, 1000, 2000, 4000] as const;

function ratio(value: number, previous: number | undefined): string {
  return previous === undefined ? "-" : `${(value / previous).toFixed(2)}x`;
}

function plain(n: number): string {
  let s = "";
  for (let i = 0; i < n; i++) {
    s += `<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Paragraph ${i} </w:t></w:r><w:r><w:t>with two runs.</w:t></w:r></w:p>`;
  }
  return s + LETTER_SECT_PR;
}

function rich(n: number): string {
  let s = "";
  for (let i = 0; i < n; i++) {
    const list =
      i % 3 === 0
        ? '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
        : "";
    s +=
      `<w:p w14:paraId="${(0x1000000 + i).toString(16).toUpperCase()}"><w:pPr><w:pStyle w:val="BenchLeaf"/>${list}<w:spacing w:after="${i % 7}0"/></w:pPr>` +
      `<w:bookmarkStart w:id="${i}" w:name="bm${i}"/>` +
      `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Item ${i} </w:t></w:r>` +
      `<w:r><w:t>text</w:t></w:r><w:bookmarkEnd w:id="${i}"/></w:p>`;
  }
  return s + LETTER_SECT_PR;
}

const styleChain =
  '<w:style w:type="paragraph" w:styleId="BenchBase"><w:name w:val="Bench Base"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="22"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="BenchMiddle"><w:name w:val="Bench Middle"/><w:basedOn w:val="BenchBase"/><w:pPr><w:spacing w:after="80"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="BenchLeaf"><w:name w:val="Bench Leaf"/><w:basedOn w:val="BenchMiddle"/><w:rPr><w:b/></w:rPr></w:style>';

function measure(build: (n: number) => string, n: number): Timings {
  const bytes = makeStyledNumberedDocx(build(n), styleChain);
  const t0 = performance.now();
  const opened = importDocx(bytes);
  const t1 = performance.now();
  const state = editorStateForSession(opened);
  const t2 = performance.now();
  const pos = Math.floor(state.doc.content.size / 2);
  const $pos = state.doc.resolve(pos);
  const sel = TextSelection.near($pos);
  let s = state.apply(state.tr.setSelection(sel));
  const t3 = performance.now();
  for (let k = 0; k < 5; k++) s = s.apply(s.tr.insertText("x"));
  const t4 = performance.now();
  exportDocx(s.doc, opened.session);
  const t5 = performance.now();
  return {
    import: t1 - t0,
    state: t2 - t1,
    edit: (t4 - t3) / 5,
    export: t5 - t4,
  };
}

const NOTE_PARAGRAPHS = 1000;
const FOOTNOTES = 280;
const ENDNOTES = 20;
const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/**
 * The paragraphs of the notes case: a footnote reference on every third paragraph until all are
 * placed and an endnote reference on every fiftieth, or the same paragraphs referring to nothing
 */
function notesBody(withReferences: boolean): string {
  let s = "";
  let footnotes = 0;
  let endnotes = 0;
  for (let i = 0; i < NOTE_PARAGRAPHS; i++) {
    let references = "";
    if (withReferences && i % 3 === 0 && footnotes < FOOTNOTES) {
      footnotes += 1;
      references += `<w:r><w:footnoteReference w:id="${footnotes}"/></w:r>`;
    }
    if (withReferences && i % 50 === 25 && endnotes < ENDNOTES) {
      endnotes += 1;
      references += `<w:r><w:endnoteReference w:id="${endnotes}"/></w:r>`;
    }
    s += `<w:p><w:r><w:t xml:space="preserve">Paragraph ${i} </w:t></w:r>${references}<w:r><w:t>with two runs.</w:t></w:r></w:p>`;
  }
  return s + LETTER_SECT_PR;
}

/** A notes part of `count` notes, each with a bold run, and every fourth with a second paragraph */
function notesPart(kind: "footnote" | "endnote", count: number): string {
  let s = `<w:${kind}s ${W_NS}><w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>`;
  for (let id = 1; id <= count; id++) {
    const second =
      id % 4 === 0
        ? "<w:p><w:r><w:t>And a second paragraph.</w:t></w:r></w:p>"
        : "";
    s +=
      `<w:${kind} w:id="${id}"><w:p><w:r><w:${kind}Ref/></w:r>` +
      `<w:r><w:t xml:space="preserve"> Note ${id} says </w:t></w:r>` +
      `<w:r><w:rPr><w:b/></w:rPr><w:t>something bold</w:t></w:r></w:p>${second}</w:${kind}>`;
  }
  return `${s}</w:${kind}s>`;
}

function notesDocx(withNotes: boolean): Uint8Array {
  if (!withNotes) return makeDocx(notesBody(false));
  const parts = unzipSync(makeNotesDocx(notesBody(true)));
  const encoder = new TextEncoder();
  parts["word/footnotes.xml"] = encoder.encode(
    notesPart("footnote", FOOTNOTES)
  );
  parts["word/endnotes.xml"] = encoder.encode(notesPart("endnote", ENDNOTES));
  return zipSync(parts);
}

/** The middle of `times` runs, which a pass as short as a layout needs to be read at all */
function median(run: () => void, times = 21): number {
  const taken: number[] = [];
  for (let k = 0; k < times; k++) {
    const t = performance.now();
    run();
    taken.push(performance.now() - t);
  }
  taken.sort((a, b) => a - b);
  return taken[Math.floor(times / 2)] ?? 0;
}

const NOTE_KEYS = [
  "import",
  "state",
  "markup",
  "demands",
  "demandsAgain",
  "layout",
] as const;

type NoteTimings = Record<(typeof NOTE_KEYS)[number], number>;

/**
 * Opens the document, draws every note's markup, asks the footnote source about every block twice
 * as two layout passes would, and lays synthetic 20px blocks out with the band the editor hands
 * over, which is none for a document referring to no footnote.
 */
function measureNotes(withNotes: boolean): NoteTimings {
  const bytes = notesDocx(withNotes);
  const t0 = performance.now();
  const opened = importDocx(bytes);
  const t1 = performance.now();
  const state = editorStateForSession(opened);
  const t2 = performance.now();
  const { footnotes, endnotes } = noteProjection.read(state);
  for (const row of [...footnotes.values(), ...endnotes]) {
    storyMarkup(row.story, {
      fontFallbacks: DEFAULT_FONT_FALLBACKS,
      nodeSpecs: noteNodeSpecs(() => row.label),
    });
  }
  const t3 = performance.now();

  const view = new EditorView(document.createElement("div"), { state });
  const targets: MeasureTarget[] = [];
  state.doc.forEach((node, pos) => {
    const dom = view.nodeDOM(pos);
    if (dom instanceof HTMLElement) {
      targets.push({
        view,
        node,
        pos,
        dom,
        sheetY: (y) => y,
        top: 0,
        scale: 1,
      });
    }
  });
  const t4 = performance.now();
  const blocks: MeasuredBlock[] = targets.map((target) => ({
    pos: target.pos,
    gap: 0,
    height: 20,
    breakBefore: false,
    breakAfter: false,
    candidates: [],
    minFirstPiece: 20,
    keepWithNext: false,
    demands: footnoteDemands.demandsIn(target),
  }));
  const t5 = performance.now();
  const demandsAgain = median(() => {
    for (const target of targets) footnoteDemands.demandsIn(target);
  });
  view.destroy();

  const sections = sectionPixels(sectionsOf(state.doc));
  const bands =
    footnotes.size === 0
      ? undefined
      : new Map<string, DemandBand>([
          [
            FOOTNOTE_BAND,
            {
              overhead: 16,
              heights: new Map([...footnotes.keys()].map((key) => [key, 36])),
            },
          ],
        ]);
  const layout = median(() => pageLayout({ blocks, sections, bands }));
  return {
    import: t1 - t0,
    state: t2 - t1,
    markup: t3 - t2,
    demands: t5 - t4,
    demandsAgain,
    layout,
  };
}

describe("notes scaling", () => {
  it(`notes: open, draw ${FOOTNOTES + ENDNOTES} notes, ask for room, and lay ${NOTE_PARAGRAPHS} paragraphs out`, () => {
    measureNotes(false);
    measureNotes(true);
    for (const withNotes of [false, true]) {
      const timings = measureNotes(withNotes);
      console.log(
        withNotes
          ? `notes n=${NOTE_PARAGRAPHS} footnotes=${FOOTNOTES} endnotes=${ENDNOTES}`
          : `no notes n=${NOTE_PARAGRAPHS}`
      );
      for (const key of NOTE_KEYS) {
        console.log(`  ${key}=${timings[key].toFixed(2)}ms`);
      }
    }
  }, 900_000);
});

describe("document scaling", () => {
  for (const [name, build] of [
    ["plain", plain],
    ["rich", rich],
  ] as const) {
    it(`${name}: import, state, keystroke, export by paragraph count`, () => {
      // JIT and module warm-up land on whichever size runs first, so one pass is thrown away
      measure(build, sizes[0]);
      let previous: Timings | undefined;
      for (const n of sizes) {
        const timings = measure(build, n);
        console.log(`${name} n=${n}`);
        for (const key of ["import", "state", "edit", "export"] as const) {
          console.log(
            `  ${key}=${timings[key].toFixed(1)}ms ratio=${ratio(timings[key], previous?.[key])}`
          );
        }
        previous = timings;
      }
    }, 900_000);
  }
});
