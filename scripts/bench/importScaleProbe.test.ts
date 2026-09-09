// @vitest-environment jsdom

import { TextSelection } from "prosemirror-state";
import { describe, it } from "vitest";
import {
  LETTER_SECT_PR,
  makeStyledNumberedDocx,
} from "../../src/__testing__/docx";
import { exportDocx } from "../../src/docx/exportDocx";
import { importDocx } from "../../src/docx/importDocx";
import { editorStateForSession } from "../../src/editor/createEditor";

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

describe("document scaling", () => {
  for (const [name, build] of [
    ["plain", plain],
    ["rich", rich],
  ] as const) {
    it(`${name}: import, state, keystroke, export by paragraph count`, () => {
      let previous: Timings | undefined;
      for (const n of sizes) {
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
        const timings = {
          import: t1 - t0,
          state: t2 - t1,
          edit: (t4 - t3) / 5,
          export: t5 - t4,
        };
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
