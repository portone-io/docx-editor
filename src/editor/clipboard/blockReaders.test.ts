// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import type { Node as PMNode } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { describe, expect, it } from "vitest";
import { makeNumberedDocx } from "../../__testing__/docx";
import { importDocx } from "../../docx/importDocx";
import { toRunFormat } from "../../model/format";
import { newListsOf } from "../../numbering/listRegistry";
import { templateList } from "../../numbering/listTemplate";
import { editorClassNames } from "../../styles/classNames";
import { listRefOf } from "../commands/listCommands";
import { createEditorView, editorStateForSession } from "../createEditor";
import { INTERNAL_TOKEN_ATTRIBUTE } from "./internalChannel";
import { detectHtmlSource } from "./source";

const htmlDir = join(
  dirname(fileURLToPath(import.meta.url)),
  "__testing__",
  "html"
);

/** One of the clipboard shapes under `__testing__/html`, as the browser would hand it over */
function fixture(name: string): string {
  return readFileSync(join(htmlDir, name), "utf8");
}

function markup(html: string): DocumentFragment {
  const template = document.createElement("template");
  template.innerHTML = html;
  return template.content;
}

function openEditor(): EditorView {
  const { doc, session } = importDocx(
    makeNumberedDocx(
      '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>'
    )
  );
  const view = createEditorView({
    mount: document.createElement("div"),
    state: editorStateForSession({ doc, session }),
    onStateChange: () => {},
  });
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7))
  );
  return view;
}

function paste(view: EditorView, html: string): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => (type === "text/html" ? html : "") },
  });
  view.dom.dispatchEvent(event);
}

/** Everything a paste of this markup put into a document holding one paragraph */
function pasted(html: string): PMNode {
  const view = openEditor();
  paste(view, html);
  const { doc } = view.state;
  view.destroy();
  return doc;
}

function nodesOfType(doc: PMNode, type: string): PMNode[] {
  const found: PMNode[] = [];
  doc.descendants((node) => {
    if (node.type.name === type) found.push(node);
    return true;
  });
  return found;
}

/** The text of every cell of a table, row by row */
function cellText(table: PMNode): string[][] {
  return table.children.map((row) =>
    row.children.map((cell) => cell.textContent)
  );
}

function runFormatOf(doc: PMNode, text: string) {
  let found = null;
  doc.descendants((node) => {
    if (!node.isText || node.text !== text) return true;
    found = toRunFormat(
      node.marks.find((mark) => mark.type.name === "run")?.attrs.format
    );
    return false;
  });
  return found;
}

describe("reading HTML written by another application", () => {
  it("identifies the source from the captured HTML", () => {
    expect(detectHtmlSource(markup(fixture("word-list-table.html")))).toBe(
      "word"
    );
    expect(
      detectHtmlSource(markup(fixture("google-docs-formatting.html")))
    ).toBe("google-docs");
    expect(
      detectHtmlSource(markup(fixture("libreoffice-paragraphs.html")))
    ).toBe("libreoffice");
    expect(
      detectHtmlSource(
        markup(
          `<p class="${editorClassNames.paragraph}" ${INTERNAL_TOKEN_ATTRIBUTE}="copy-1">own</p>`
        )
      )
    ).toBe("editor");
    expect(detectHtmlSource(markup("<p>plain</p>"))).toBe("unknown");
  });

  it("reads a Word clipboard table as a table with one paragraph per cell", () => {
    const doc = pasted(fixture("word-list-table.html"));

    const tables = nodesOfType(doc, "table");
    expect(tables).toHaveLength(1);
    const table = tables[0];
    if (!table) throw new Error("no table pasted");
    expect(cellText(table)).toEqual([
      ["Crate", "Weight"],
      ["Ballast", "Twelve kilograms"],
    ]);
    for (const row of table.children) {
      for (const cell of row.children) {
        expect(cell.childCount).toBe(1);
        expect(cell.firstChild?.type.name).toBe("paragraph");
      }
    }
  });

  it("reads Word mso-list paragraphs as a list", () => {
    const doc = pasted(fixture("word-list-table.html"));

    const items = nodesOfType(doc, "paragraph").filter((block) =>
      ["Spare rope", "Weather chart"].includes(block.textContent)
    );
    expect(items).toHaveLength(2);
    const refs = items.map((item) => listRefOf(item));
    expect(refs[0]).toEqual({ numId: refs[1]?.numId, ilvl: 0 });
    expect(refs[0]?.numId).toEqual(expect.any(Number));
    // Word draws the bullet into the paragraph, and the list numbers itself here
    expect(doc.textContent).not.toContain("\u00b7");
  });

  it("keeps two Word lists apart and counts the one whose marker counts", () => {
    const doc = pasted(
      '<meta name=Generator content="Microsoft Word 15">' +
        "<p class=MsoListParagraphCxSpFirst style='mso-list:l0 level1 lfo1'>" +
        "<![if !supportLists]><span style='mso-list:Ignore'>1.</span><![endif]>First</p>" +
        "<p class=MsoListParagraphCxSpLast style='mso-list:l0 level1 lfo1'>" +
        "<![if !supportLists]><span style='mso-list:Ignore'>2.</span><![endif]>Second</p>" +
        "<p class=MsoNormal style='mso-list:l1 level2 lfo2'>" +
        "<span style='mso-list:Ignore'>\u00b7</span>Aside</p>"
    );

    const items = nodesOfType(doc, "paragraph").filter((block) =>
      ["First", "Second", "Aside"].includes(block.textContent)
    );
    expect(items.map((item) => listRefOf(item)?.ilvl)).toEqual([0, 0, 1]);
    const [first, second, aside] = items.map(
      (item) => listRefOf(item)?.numId ?? null
    );
    expect(first).toBe(second);
    expect(aside).not.toBe(first);
    const registered = newListsOf(doc.attrs.newLists);
    expect(registered.get(first ?? -1)).toEqual(templateList("numbered"));
    expect(registered.get(aside ?? -1)).toEqual(templateList("bullet"));
  });

  it("gives a cell one paragraph for each block it holds", () => {
    const doc = pasted(
      "<table><tr><td><p>first</p><p>second</p></td><td>only</td></tr></table>"
    );

    const cells = nodesOfType(doc, "tableCell");
    expect(cells).toHaveLength(2);
    expect(cells[0]?.children.map((block) => block.textContent)).toEqual([
      "first",
      "second",
    ]);
    expect(cells[1]?.childCount).toBe(1);
  });

  it("keeps the columns and rows a pasted cell reaches across", () => {
    const doc = pasted(
      "<table>" +
        '<tr><td colspan="2">wide</td><td rowspan="2">tall</td></tr>' +
        "<tr><td>a</td><td>b</td></tr>" +
        "</table>"
    );

    const table = nodesOfType(doc, "table")[0];
    expect(table?.attrs.gridCols).toHaveLength(3);
    expect(cellText(table ?? doc)).toEqual([
      ["wide", "tall"],
      ["a", "b"],
    ]);
    const cells = nodesOfType(doc, "tableCell");
    expect(
      cells.map((cell) => [cell.attrs.colspan, cell.attrs.rowspan])
    ).toEqual([
      [2, 1],
      [1, 2],
      [1, 1],
      [1, 1],
    ]);
  });

  it("counts a Word list whose marker it could not find", () => {
    const doc = pasted(
      '<meta name=Generator content="Microsoft Word 15">' +
        "<p class=MsoNormal style='mso-list:l0 level1 lfo1'>" +
        "<span style='font-size:11.0pt'>1.\u00a0\u00a0</span>Rope</p>" +
        "<p class=MsoNormal style='mso-list:l0 level1 lfo1'>" +
        "<span style='font-size:11.0pt'>2.\u00a0\u00a0</span>Chart</p>" +
        "<p class=MsoNormal style='mso-list:l1 level1 lfo2'>" +
        "<span style='font-family:Symbol'>\u00b7</span>Cord</p>"
    );

    const items = nodesOfType(doc, "paragraph").filter((block) =>
      ["Rope", "Chart", "Cord"].includes(block.textContent)
    );
    expect(items).toHaveLength(3);
    const [rope, chart, cord] = items.map(
      (item) => listRefOf(item)?.numId ?? null
    );
    expect(rope).toBe(chart);
    expect(cord).not.toBe(rope);
    const registered = newListsOf(doc.attrs.newLists);
    expect(registered.get(rope ?? -1)).toEqual(templateList("numbered"));
    expect(registered.get(cord ?? -1)).toEqual(templateList("bullet"));
  });

  it("keeps a table's caption", () => {
    const doc = pasted(
      "<table><caption>Crates and weights</caption>" +
        "<tr><td>Ballast</td><td>Twelve</td></tr></table>"
    );

    const blocks = doc.children.map((block) => [
      block.type.name,
      block.textContent,
    ]);
    expect(blocks).toContainEqual(["paragraph", "Crates and weights"]);
    const caption = blocks.findIndex(
      ([, text]) => text === "Crates and weights"
    );
    expect(blocks[caption + 1]?.[0]).toBe("table");
  });

  it("keeps a table larger than the model's side limit readable", () => {
    const rows = Array.from(
      { length: 51 },
      (_unused, row) => `<tr><td>a${row}</td><td>b${row}</td></tr>`
    ).join("");
    const doc = pasted(`<table>${rows}</table>`);

    expect(nodesOfType(doc, "table")).toHaveLength(0);
    const read = nodesOfType(doc, "paragraph").map(
      (block) => block.textContent
    );
    expect(read).toHaveLength(51);
    expect(read[0]).toBe("a0\tb0");
    expect(read[50]).toBe("a50\tb50");
  });

  it("drops a row a table gives no cells", () => {
    const doc = pasted("<table><tr></tr><tr><td>a</td></tr></table>");

    const table = nodesOfType(doc, "table")[0];
    expect(table?.childCount).toBe(1);
    expect(nodesOfType(doc, "tableCell")).toHaveLength(1);
  });

  it("reads a table inside a cell as the text of that cell", () => {
    const doc = pasted(
      "<table><tr><td>outer<table><tr><td>inner</td></tr></table></td></tr></table>"
    );

    expect(nodesOfType(doc, "table")).toHaveLength(1);
    expect(nodesOfType(doc, "tableCell")[0]?.textContent).toContain("inner");
  });

  it("reads Google Docs bold and colored spans into run properties", () => {
    const doc = pasted(fixture("google-docs-formatting.html"));

    expect(runFormatOf(doc, "lantern workshop")).toMatchObject({ bold: true });
    expect(runFormatOf(doc, ", second draft")).toMatchObject({
      color: "#FF0000",
    });
    expect(runFormatOf(doc, "Signed off by the shop steward")).toMatchObject({
      underline: "single",
    });
    expect(doc.textContent).toContain("Notes from the lantern workshop");
    // The whole copy is wrapped in one `<b>`, which must not press its paragraphs into one
    expect(
      nodesOfType(doc, "paragraph").filter((block) =>
        block.textContent.includes("Signed off")
      )
    ).toHaveLength(1);
    expect(
      nodesOfType(doc, "paragraph").some(
        (block) =>
          block.textContent === "Notes from the lantern workshop, second draft"
      )
    ).toBe(true);
  });

  it("keeps LibreOffice paragraphs and breaks", () => {
    const doc = pasted(fixture("libreoffice-paragraphs.html"));

    const paragraphs = nodesOfType(doc, "paragraph").map(
      (block) => block.textContent
    );
    expect(paragraphs).toContain(
      "The kite ran out of string and drifted over the marsh"
    );
    expect(paragraphs).toContain(
      "A second paragraph, kept apart from the first"
    );
    expect(nodesOfType(doc, "hardBreak")).toHaveLength(1);
    expect(
      runFormatOf(doc, "A second paragraph, kept apart from the first")
    ).toMatchObject({ italic: true });
  });
});
