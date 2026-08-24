// @vitest-environment jsdom
import { unzipSync } from "fflate";
import { TextSelection } from "prosemirror-state";
import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  decode,
  makeDocx,
  makeNumberedDocx,
  makeStyledDocx,
} from "../__testing__/docx";
import { renderInto } from "../__testing__/react";
import {
  DocxEditor,
  type DocxEditorHandle,
  type DocxEditorMode,
} from "../DocxEditor";
import { IMAGE_FILE_ACCEPT } from "../editor/commands";
import { lockSelection } from "../editor/commands/lockCommands";
import { editorClassNames } from "../styles/classNames";
import type { DocxEditorPresets } from "./presets";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => {
  host.remove();
});

const PARAGRAPH = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>'
);

/** A background color that an older document recorded with w:highlight */
const HIGHLIGHTED = makeDocx(
  "<w:p><w:r>" +
    '<w:rPr><w:highlight w:val="yellow"/></w:rPr>' +
    '<w:t xml:space="preserve">source</w:t>' +
    "</w:r></w:p>"
);

/** The same paragraph, but in a document that has somewhere (numbering.xml) to write a new list definition */
const NUMBERABLE = makeNumberedDocx(
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>'
);

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

const WITH_TABLE = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>' +
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr>${cellXml("Left")}${cellXml("Right")}</w:tr>` +
    "</w:tbl>"
);

const render = (element: ReactNode) => renderInto(host, element);

interface Mounted {
  handle: DocxEditorHandle;
  unmount: () => void;
}

function mount(
  bytes: Uint8Array,
  props: { mode?: DocxEditorMode; presets?: DocxEditorPresets } = {}
) {
  const box: { current: DocxEditorHandle | null } = { current: null };
  const unmount = render(
    <DocxEditor
      document={bytes}
      ref={box}
      renderImportError={() => null}
      {...props}
    />
  );
  const handle = box.current;
  if (!handle) throw new Error("the ref was not attached");
  // jsdom draws nothing, so the spot a right click lands on is answered here: the caret the
  // state already holds, which is what a click on the current selection lands on
  handle.view.posAtCoords = () => ({
    pos: handle.view.state.selection.head,
    inside: -1,
  });
  return { handle, unmount } satisfies Mounted;
}

function button(label: string): HTMLButtonElement {
  const found = host.querySelector(`button[aria-label="${label}"]`);
  if (!(found instanceof HTMLButtonElement)) {
    throw new Error(`button not found: ${label}`);
  }
  return found;
}

function selectBox(label: string): HTMLSelectElement {
  const found = host.querySelector(`select[aria-label="${label}"]`);
  if (!(found instanceof HTMLSelectElement)) {
    throw new Error(`select not found: ${label}`);
  }
  return found;
}

function click(label: string): void {
  act(() => button(label).click());
}

/** A row of an open menu, found by the text written on it */
function menuItem(text: string): HTMLButtonElement {
  const found = Array.from(host.querySelectorAll("button[role^='menuitem']"))
    .filter(
      (element): element is HTMLButtonElement =>
        element instanceof HTMLButtonElement
    )
    .find((element) => element.textContent === text);
  if (!found) throw new Error(`menu item not found: ${text}`);
  return found;
}

/** Every control on the toolbar, named by its label, in the order they are drawn */
function toolbarLabels(): (string | null)[] {
  return Array.from(
    host.querySelectorAll('[role="toolbar"] button, [role="toolbar"] select')
  )
    .filter((element) => element.hasAttribute("aria-label"))
    .map((element) => element.getAttribute("aria-label"));
}

/** Which icon a toolbar button is wearing, told apart by the class the icon set puts on it */
function buttonIcon(label: string): string | null {
  return button(label).querySelector("svg")?.getAttribute("class") ?? null;
}

/** The icon of the entry the open alignment menu marks. It comes after the check mark in the row */
function markedAlignIcon(): string | null {
  const marked = host.querySelector(
    '[role="menuitemradio"][aria-checked="true"]'
  );
  const icons = Array.from(marked?.querySelectorAll("svg") ?? []);
  return icons[icons.length - 1]?.getAttribute("class") ?? null;
}

/** Entries without a label are found by the text written on them */
function clickText(text: string): void {
  const found = Array.from(host.querySelectorAll("button")).find(
    (element) => element.textContent === text
  );
  if (!found) throw new Error(`menu item not found: ${text}`);
  act(() => found.click());
}

function documentXml(handle: DocxEditorHandle): string {
  return decode(unzipSync(handle.exportBytes())["word/document.xml"]);
}

/** The formatting XML the first paragraph carries. An empty string when there is none */
function firstPPr(handle: DocxEditorHandle): string {
  const pPr: unknown = handle.view.state.doc.child(0).attrs.pPr;
  return typeof pPr === "string" ? pPr : "";
}

/** Selects all the text in the first paragraph */
function selectFirstParagraph(handle: DocxEditorHandle): void {
  act(() => {
    const { view } = handle;
    const paragraph = view.state.doc.child(0);
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, 1, paragraph.nodeSize - 1)
      )
    );
  });
}

describe("the built in toolbar", () => {
  it("shows by default and disappears when readOnly or turned off", () => {
    const shown = mount(PARAGRAPH);
    expect(host.querySelector('[role="toolbar"]')).not.toBeNull();
    shown.unmount();

    const off = mount(PARAGRAPH, {
      mode: { kind: "edit", toolbar: false },
    });
    expect(host.querySelector('[role="toolbar"]')).toBeNull();
    off.unmount();

    const box: { current: DocxEditorHandle | null } = { current: null };
    const unmount = render(
      <DocxEditor
        document={PARAGRAPH}
        ref={box}
        mode={{ kind: "readOnly" }}
        renderImportError={() => null}
      />
    );
    expect(host.querySelector('[role="toolbar"]')).toBeNull();
    unmount();
  });

  /**
   * The four letter shapes follow the font the letters are set in, and come before their
   * color: the order the controls stand in is what tells the groups apart on screen
   */
  it("draws the bold group straight after the font controls", () => {
    const { unmount } = mount(PARAGRAPH);
    const labels = toolbarLabels();

    expect(
      labels.slice(labels.indexOf("Style"), labels.indexOf("Text color") + 1)
    ).toEqual([
      "Style",
      "Font",
      "Font size",
      "Bold",
      "Italic",
      "Underline",
      "Strikethrough",
      "Text color",
    ]);
    unmount();
  });

  it("clicking bold applies it to the selected text and presses the button", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    expect(button("Bold").getAttribute("aria-pressed")).toBe("false");

    click("Bold");
    expect(documentXml(handle)).toContain("<w:b/>");
    expect(button("Bold").getAttribute("aria-pressed")).toBe("true");

    click("Bold");
    expect(documentXml(handle)).not.toContain("<w:b/>");
    unmount();
  });

  it("undo is only clickable once there is an edit to undo", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    expect(button("Undo").disabled).toBe(true);

    selectFirstParagraph(handle);
    click("Bold");
    expect(button("Undo").disabled).toBe(false);

    click("Undo");
    expect(documentXml(handle)).not.toContain("<w:b/>");
    unmount();
  });

  it("the indent buttons move a plain paragraph half an inch at a time", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    // A paragraph sitting at the left margin has nothing to give up
    expect(button("Decrease indent").disabled).toBe(true);
    expect(button("Increase indent").disabled).toBe(false);

    click("Increase indent");
    expect(documentXml(handle)).toContain('<w:ind w:left="720"/>');
    expect(button("Decrease indent").disabled).toBe(false);

    click("Decrease indent");
    expect(documentXml(handle)).not.toContain("<w:ind");
    unmount();
  });

  it("in a list the same buttons move the item's level", () => {
    const { handle, unmount } = mount(NUMBERABLE);
    selectFirstParagraph(handle);
    click("Numbered list");
    // The first level is as far left as an item goes
    expect(button("Decrease indent").disabled).toBe(true);

    click("Increase indent");
    expect(firstPPr(handle)).toContain('<w:ilvl w:val="1"/>');
    // The level's own indentation is what moves the item, so the paragraph records none
    expect(firstPPr(handle)).not.toContain("<w:ind");
    expect(button("Decrease indent").disabled).toBe(false);
    unmount();
  });

  /**
   * A new list's definition has to be written into numbering.xml to survive export.
   * If the button were clickable in a document without that part, it would create an
   * edit that becomes a list on screen but is blocked on export.
   */
  it("the list buttons are disabled in a document with no numbering.xml", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);

    expect(button("Numbered list").disabled).toBe(true);
    expect(button("Bulleted list").disabled).toBe(true);
    unmount();
  });

  it("the list buttons toggle on and off and switch between kinds", () => {
    const { handle, unmount } = mount(NUMBERABLE);
    const listed = () => firstPPr(handle);

    selectFirstParagraph(handle);
    expect(button("Numbered list").getAttribute("aria-pressed")).toBe("false");

    click("Numbered list");
    expect(button("Numbered list").getAttribute("aria-pressed")).toBe("true");
    expect(listed()).toContain("<w:numPr>");

    // Clicking a different kind switches to that kind
    click("Bulleted list");
    expect(button("Bulleted list").getAttribute("aria-pressed")).toBe("true");
    expect(button("Numbered list").getAttribute("aria-pressed")).toBe("false");

    // Clicking the same kind again leaves the list
    click("Bulleted list");
    expect(button("Bulleted list").getAttribute("aria-pressed")).toBe("false");
    expect(listed()).not.toContain("<w:numPr>");
    unmount();
  });

  it("the alignment menu applies the alignment picked and marks the one in force", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    click("Alignment");
    // A paragraph where nothing declares an alignment is left aligned
    expect(menuItem("Align left").getAttribute("aria-checked")).toBe("true");
    expect(menuItem("Justify").getAttribute("aria-checked")).toBe("false");

    clickText("Justify");
    // Justify is written out to the document as both
    expect(documentXml(handle)).toContain('<w:jc w:val="both"/>');
    // The panel closes once a choice is made
    expect(host.querySelector('[role="menuitemradio"]')).toBeNull();

    click("Alignment");
    expect(menuItem("Justify").getAttribute("aria-checked")).toBe("true");
    expect(menuItem("Align left").getAttribute("aria-checked")).toBe("false");

    clickText("Align center");
    expect(documentXml(handle)).toContain('<w:jc w:val="center"/>');
    unmount();
  });

  it("the one alignment button wears the current alignment's icon", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);

    click("Alignment");
    const leftAligned = buttonIcon("Alignment");
    // The button wears the icon of the entry the menu marks
    expect(leftAligned).toBe(markedAlignIcon());

    clickText("Align right");
    click("Alignment");
    expect(buttonIcon("Alignment")).toBe(markedAlignIcon());
    expect(buttonIcon("Alignment")).not.toBe(leftAligned);

    // The four buttons that used to sit in a row are gone
    const labels = Array.from(host.querySelectorAll("button[aria-label]")).map(
      (element) => element.getAttribute("aria-label")
    );
    expect(labels).not.toContain("Align center");
    unmount();
  });

  it("the spacing menu marks the line spacing in force and applies the one picked", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    click("Line and paragraph spacing");

    // A paragraph where nobody declared a spacing is drawn on a single line
    expect(menuItem("Single").getAttribute("aria-checked")).toBe("true");
    expect(menuItem("1.5").getAttribute("aria-checked")).toBe("false");

    clickText("1.5");
    expect(documentXml(handle)).toContain(
      '<w:spacing w:line="360" w:lineRule="auto"/>'
    );
    // The panel closes once a choice is made
    expect(host.querySelector('[role="menuitemradio"]')).toBeNull();

    click("Line and paragraph spacing");
    expect(menuItem("1.5").getAttribute("aria-checked")).toBe("true");
    expect(menuItem("Single").getAttribute("aria-checked")).toBe("false");
    unmount();
  });

  it("the spacing menu offers the four presets and nothing else", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    click("Line and paragraph spacing");

    const entries = Array.from(
      host.querySelectorAll("button[role^='menuitem']")
    ).map((element) => element.textContent);
    expect(entries).toEqual(["Single", "1.15", "1.5", "Double"]);
    unmount();
  });

  it("inserts a table at the size picked from the grid", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    click("Insert table");
    click("2 by 3 table");

    const inserted = handle.view.state.doc.child(1);
    expect(inserted.type.name).toBe("table");
    expect(inserted.childCount).toBe(2);
    expect(inserted.child(0).childCount).toBe(3);
    // The panel closes once a choice is made
    expect(host.querySelector('button[aria-label="2 by 3 table"]')).toBeNull();
    unmount();
  });

  it("the image button opens a picker for the kinds we can write back", () => {
    const { unmount } = mount(PARAGRAPH);
    expect(button("Insert image").disabled).toBe(false);

    const picker = host.querySelector('input[type="file"]');
    if (!(picker instanceof HTMLInputElement)) {
      throw new Error("no file picker");
    }
    expect(picker.accept).toBe(IMAGE_FILE_ACCEPT);
    // The picker never shows itself; the toolbar button is the only thing on screen
    expect(picker.hidden).toBe(true);

    let opened = 0;
    picker.click = () => {
      opened += 1;
    };
    click("Insert image");
    expect(opened).toBe(1);
    unmount();
  });

  /** A document holding nothing the editor can model, so an image has nowhere to go */
  const ONLY_PRESERVED = makeDocx("<w:sdt><w:sdtContent/></w:sdt>");

  it("blocks inserting an image where no inline content can stand", () => {
    const { unmount } = mount(ONLY_PRESERVED);
    expect(button("Insert image").disabled).toBe(true);
    unmount();
  });

  it("blocks inserting a table inside a table", () => {
    const { handle, unmount } = mount(WITH_TABLE);
    act(() => {
      const { view } = handle;
      const at = view.state.doc.child(0).nodeSize + 4;
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.near(view.state.doc.resolve(at))
        )
      );
    });
    expect(button("Insert table").disabled).toBe(true);
    unmount();
  });
});

/** Two runs at different sizes */
const MIXED_SIZES = makeDocx(
  "<w:p>" +
    '<w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">ab</w:t></w:r>' +
    '<w:r><w:rPr><w:sz w:val="48"/></w:rPr><w:t xml:space="preserve">cd</w:t></w:r>' +
    "</w:p>"
);

/** A document that declares only a 12pt document default size and nothing on the runs */
const DEFAULT_12PT = makeDocx(
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>',
  '<w:rPr><w:sz w:val="24"/></w:rPr>'
);

describe("showing the font size", () => {
  const fontSize = () => selectBox("Font size");

  it("shows that number where a size is written down", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    act(() => {
      const select = fontSize();
      select.value = "14";
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });

    expect(fontSize().value).toBe("14");
    unmount();
  });

  it("shows the document default size where nobody wrote one down", () => {
    const shown = mount(DEFAULT_12PT);
    selectFirstParagraph(shown.handle);
    expect(fontSize().value).toBe("12");
    shown.unmount();

    // When the document declares no default size, the fallback is 10pt, the same as Word
    const fallback = mount(PARAGRAPH);
    selectFirstParagraph(fallback.handle);
    expect(fontSize().value).toBe("10");
    fallback.unmount();
  });

  it("the Default entry is something to pick, not a displayed value", () => {
    const { handle, unmount } = mount(DEFAULT_12PT);
    selectFirstParagraph(handle);
    // The Default entry stays in place even while the default size is being shown
    expect(fontSize().value).toBe("12");
    const labels = Array.from(fontSize().options).map((one) => one.textContent);
    expect(labels).toContain("Default");
    unmount();
  });
});

/** A document whose font is not among the presets */
const DOCUMENT_FONT = makeDocx(
  "<w:p><w:r>" +
    '<w:rPr><w:rFonts w:ascii="Custom Serif" w:hAnsi="Custom Serif" ' +
    'w:eastAsia="Custom Serif"/></w:rPr>' +
    '<w:t xml:space="preserve">source</w:t>' +
    "</w:r></w:p>"
);

/** A Japanese run: a Latin font for the letters, a mincho for everything else */
const JAPANESE_RUN = makeDocx(
  "<w:p><w:r>" +
    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="MS Mincho" ' +
    'w:hAnsi="Times New Roman"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr>' +
    '<w:t xml:space="preserve">第一条</w:t>' +
    "</w:r></w:p>"
);

/** Two runs with different fonts */
const MIXED_FONTS = makeDocx(
  "<w:p>" +
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' +
    '<w:t xml:space="preserve">ab</w:t></w:r>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Batang" w:hAnsi="Batang"/></w:rPr>' +
    '<w:t xml:space="preserve">cd</w:t></w:r>' +
    "</w:p>"
);

describe("the font picker", () => {
  const fontSelect = () => selectBox("Font");

  function choose(value: string): void {
    act(() => {
      const select = fontSelect();
      select.value = value;
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });
  }

  function labels(): (string | null)[] {
    return Array.from(fontSelect().options).map((one) => one.textContent);
  }

  it("the picked font is written to the document and shows in the select", () => {
    const { handle, unmount } = mount(PARAGRAPH, {
      presets: { fonts: ["Batang", "Arial"] },
    });
    selectFirstParagraph(handle);
    choose("Batang");

    expect(documentXml(handle)).toContain(
      '<w:rFonts w:ascii="Batang" w:hAnsi="Batang" w:eastAsia="Batang"/>'
    );
    expect(fontSelect().value).toBe("Batang");
    unmount();
  });

  it("a Latin font picked for a Japanese run leaves its East Asian font alone", () => {
    const { handle, unmount } = mount(JAPANESE_RUN);
    selectFirstParagraph(handle);
    choose("Arial");

    expect(documentXml(handle)).toContain(
      '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho"/>'
    );
    // The picker shows the Latin font, which is the one it wrote
    expect(fontSelect().value).toBe("Arial");
    unmount();
  });

  it("the fonts the document uses show alongside the presets", () => {
    const { handle, unmount } = mount(DOCUMENT_FONT);
    selectFirstParagraph(handle);

    expect(labels()).toContain("Georgia");
    expect(labels()).toContain("Custom Serif");
    expect(fontSelect().value).toBe("Custom Serif");
    unmount();
  });

  it("shows the font that actually gets rendered when nothing is specified", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    // When the document declares no default font, this is the font the fallback stack lays down
    expect(fontSelect().value).toBe("Arial");
    expect(labels()).toContain("Default");
    unmount();
  });

  it("offers the fonts handed in instead of the built-in ones", () => {
    const { handle, unmount } = mount(PARAGRAPH, {
      presets: { fonts: ["MS Mincho", "SimSun"] },
    });
    selectFirstParagraph(handle);

    expect(labels()).toContain("MS Mincho");
    expect(labels()).toContain("SimSun");
    expect(labels()).not.toContain("Georgia");
    unmount();
  });

  it("offers a font the document uses even when the list handed in leaves it out", () => {
    const { handle, unmount } = mount(DOCUMENT_FONT, {
      presets: { fonts: ["MS Mincho"] },
    });
    selectFirstParagraph(handle);

    expect(labels()).toContain("Custom Serif");
    expect(fontSelect().value).toBe("Custom Serif");
    unmount();
  });

  it("the Default entry withdraws the font setting", () => {
    const { handle, unmount } = mount(DOCUMENT_FONT);
    selectFirstParagraph(handle);
    choose(";default");

    expect(documentXml(handle)).not.toContain("<w:rFonts");
    unmount();
  });
});

describe("a select over a selection of mixed values", () => {
  it.each([
    ["Font size", MIXED_SIZES],
    ["Font", MIXED_FONTS],
  ])("%s: is left blank", (label, bytes) => {
    const { handle, unmount } = mount(bytes);
    selectFirstParagraph(handle);

    expect(selectBox(label).value).toBe("");
    unmount();
  });
});

/** A document defining primary and non-primary paragraph styles */
const STYLED = makeStyledDocx(
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
    '<w:r><w:t xml:space="preserve">heading</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>' +
    '<w:r><w:t xml:space="preserve">quotation</w:t></w:r></w:p>',
  '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
    '<w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1">' +
    '<w:name w:val="heading 1"/><w:qFormat/><w:rPr><w:b/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Quote">' +
    '<w:name w:val="Quote"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Index1">' +
    '<w:name w:val="index 1"/><w:semiHidden/></w:style>' +
    '<w:style w:type="table" w:styleId="TableGrid">' +
    '<w:name w:val="Table Grid"/></w:style>'
);

/** The same styles, with the second paragraph pointing at the one Word hides */
const HIDDEN_STYLED = makeStyledDocx(
  '<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:pStyle w:val="Index1"/></w:pPr>' +
    '<w:r><w:t xml:space="preserve">index</w:t></w:r></w:p>',
  '<w:style w:type="paragraph" w:styleId="Normal" w:default="1">' +
    '<w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1">' +
    '<w:name w:val="heading 1"/><w:qFormat/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Index1">' +
    '<w:name w:val="index 1"/><w:semiHidden/></w:style>'
);

describe("the paragraph style picker", () => {
  const styleSelect = () => selectBox("Style");

  /** Selects all the text in the paragraph at that index */
  function selectParagraph(handle: DocxEditorHandle, index: number): void {
    act(() => {
      const { view } = handle;
      const doc = view.state.doc;
      let start = 0;
      doc.forEach((_block, offset, at) => {
        if (at === index) start = offset + 1;
      });
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(
            doc,
            start,
            start + doc.child(index).content.size
          )
        )
      );
    });
  }

  function choose(value: string): void {
    act(() => {
      const select = styleSelect();
      select.value = value;
      select.dispatchEvent(new Event("change", { bubbles: true }));
    });
  }

  function labels(): (string | null)[] {
    return Array.from(styleSelect().options).map((one) => one.textContent);
  }

  it("offers primary paragraph styles with the default one first", () => {
    const { handle, unmount } = mount(STYLED);
    selectFirstParagraph(handle);

    expect(labels()).toEqual(["Normal", "heading 1"]);
    unmount();
  });

  it("does not offer a non-primary style even while a paragraph points at it", () => {
    const { handle, unmount } = mount(STYLED);
    selectParagraph(handle, 2);

    expect(labels()).toEqual(["", "Normal", "heading 1"]);
    expect(styleSelect().value).toBe("");
    unmount();
  });

  it("a document that defines no styles offers only Normal", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);

    expect(labels()).toEqual(["Normal"]);
    unmount();
  });

  it("the picked style is written to the document and shows in the select", () => {
    const { handle, unmount } = mount(STYLED);
    selectFirstParagraph(handle);
    expect(styleSelect().value).toBe("default");

    choose("id:Heading1");
    expect(documentXml(handle)).toContain(
      '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
    );
    expect(styleSelect().value).toBe("id:Heading1");
    unmount();
  });

  // Text typed after the style was picked carries no run of its own, so the paragraph draws it
  it("the paragraph is drawn in the character formatting of the style picked", () => {
    const { handle, unmount } = mount(STYLED);
    selectFirstParagraph(handle);
    choose("id:Heading1");

    const paragraph = host.querySelector(`p.${editorClassNames.paragraph}`);
    expect(paragraph?.getAttribute("style")).toContain("font-weight: bold");
    unmount();
  });

  it("the default entry takes the style away again", () => {
    const { handle, unmount } = mount(STYLED);
    selectFirstParagraph(handle);
    choose("id:Heading1");

    choose("default");
    expect(firstPPr(handle)).toBe("");
    unmount();
  });

  it("does not offer a hidden style even while a paragraph points at it", () => {
    const { handle, unmount } = mount(HIDDEN_STYLED);
    selectParagraph(handle, 0);
    expect(labels()).toEqual(["Normal", "heading 1"]);

    selectParagraph(handle, 1);
    expect(labels()).toEqual(["", "Normal", "heading 1"]);
    expect(styleSelect().value).toBe("");
    unmount();
  });

  // With no paragraph in the selection there is no style to show and none to pick
  it("is turned off where the selection holds no paragraph", () => {
    const { unmount } = mount(makeDocx("<w:customXml/>"));

    expect(styleSelect().disabled).toBe(true);
    expect(styleSelect().value).toBe("");
    unmount();
  });

  it("leaves it blank where the selected paragraphs point at different styles", () => {
    const { handle, unmount } = mount(STYLED);
    act(() => {
      const { view } = handle;
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(
            view.state.doc,
            1,
            view.state.doc.content.size - 1
          )
        )
      );
    });

    expect(styleSelect().value).toBe("");
    unmount();
  });
});

describe("the color palette", () => {
  const swatches = () =>
    Array.from(host.querySelectorAll('button[aria-label^="#"]'));

  const colorIndicator = (label: string) => {
    const indicator = button(label).querySelector<HTMLElement>(
      `.${editorClassNames.toolbarColorIndicator}`
    );
    if (!indicator) throw new Error(`color indicator not found: ${label}`);
    return indicator;
  };

  it("the text color is picked from a 10x8 palette", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    expect(colorIndicator("Text color").style.backgroundColor).toBe(
      "rgb(0, 0, 0)"
    );
    click("Text color");

    expect(swatches()).toHaveLength(80);
    click("#ff0000");
    expect(documentXml(handle)).toContain('<w:color w:val="FF0000"/>');
    expect(colorIndicator("Text color").style.backgroundColor).toBe(
      "rgb(255, 0, 0)"
    );
    // The panel closes once a choice is made
    expect(swatches()).toHaveLength(0);
    unmount();
  });

  it("the color currently applied is marked in the palette", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    click("Text color");
    click("#1155cc");

    click("Text color");
    expect(button("#1155cc").getAttribute("aria-selected")).toBe("true");
    expect(button("#ff0000").getAttribute("aria-selected")).toBe("false");
    unmount();
  });

  it("the None at the top of the palette withdraws the text color", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    click("Text color");
    click("#ff0000");
    expect(documentXml(handle)).toContain("<w:color");

    click("Text color");
    // The clear entry comes before the swatches
    const items = Array.from(
      host.querySelectorAll(`.${editorClassNames.popover} button`)
    );
    expect(items[0]?.textContent).toBe("None");

    clickText("None");
    expect(documentXml(handle)).not.toContain("<w:color");
    unmount();
  });

  it("the background color is picked from the same palette and withdrawn by the None in the same place", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    expect(colorIndicator("Highlight").style.backgroundColor).toBe(
      "transparent"
    );
    click("Highlight");
    const items = Array.from(
      host.querySelectorAll(`.${editorClassNames.popover} button`)
    );
    expect(items[0]?.textContent).toBe("None");
    expect(swatches()).toHaveLength(80);

    // Colors outside the 16 highlighter colors can be picked too
    click("#fff2cc");
    expect(colorIndicator("Highlight").style.backgroundColor).toBe(
      "rgb(255, 242, 204)"
    );
    expect(documentXml(handle)).toContain(
      '<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>'
    );

    click("Highlight");
    clickText("None");
    expect(documentXml(handle)).not.toContain("<w:shd");
    unmount();
  });

  it("the color of an old highlight document shows in the palette and turns into shading once painted", () => {
    const { handle, unmount } = mount(HIGHLIGHTED);
    selectFirstParagraph(handle);

    click("Highlight");
    expect(button("#ffff00").getAttribute("aria-selected")).toBe("true");

    click("#ff0000");
    const xml = documentXml(handle);
    expect(xml).not.toContain("<w:highlight");
    expect(xml).toContain(
      '<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>'
    );
    unmount();
  });
});

/**
 * Every list a built-in picker offers can be replaced from outside, and a field left out of
 * `presets` leaves that picker on the built-in list.
 */
describe("the picker lists handed in through presets", () => {
  const swatchLabels = () =>
    Array.from(host.querySelectorAll('button[aria-label^="#"]')).map(
      (element) => element.getAttribute("aria-label")
    );

  const menuLabels = () =>
    Array.from(host.querySelectorAll("button[role^='menuitem']")).map(
      (element) => element.textContent
    );

  const optionLabels = (label: string) =>
    Array.from(selectBox(label).options).map((option) => option.textContent);

  /** Puts the caret in the first cell of the table, which is the second body block */
  function caretInFirstCell(handle: DocxEditorHandle): void {
    act(() => {
      const { view } = handle;
      const at = view.state.doc.child(0).nodeSize + 4;
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.near(view.state.doc.resolve(at))
        )
      );
    });
  }

  it("draws the colors handed in instead of the built-in palette", () => {
    const { handle, unmount } = mount(PARAGRAPH, {
      presets: { colors: [["#102030", "#405060"]] },
    });
    selectFirstParagraph(handle);
    click("Text color");

    expect(swatchLabels()).toEqual(["#102030", "#405060"]);

    click("#102030");
    expect(documentXml(handle)).toContain('<w:color w:val="102030"/>');
    unmount();
  });

  it("draws those same colors in the cell pickers", () => {
    const { handle, unmount } = mount(WITH_TABLE, {
      presets: { colors: [["#102030", "#405060"]] },
    });
    caretInFirstCell(handle);
    click("Cell fill");

    expect(swatchLabels()).toEqual(["#102030", "#405060"]);
    unmount();
  });

  it("offers the sizes handed in instead of the built-in ones", () => {
    const { handle, unmount } = mount(PARAGRAPH, {
      presets: { fontSizes: [9, 40] },
    });
    selectFirstParagraph(handle);

    // 10pt is the size this document is drawn at, which is offered whatever the list holds
    expect(optionLabels("Font size")).toEqual([
      "Default",
      "10pt",
      "9pt",
      "40pt",
    ]);
    unmount();
  });

  it("offers the line spacings handed in instead of the built-in four", () => {
    const { handle, unmount } = mount(PARAGRAPH, {
      presets: {
        lineSpacings: [
          { label: "Triple", spacing: { rule: "auto", lines: 3 } },
        ],
      },
    });
    selectFirstParagraph(handle);
    click("Line and paragraph spacing");

    expect(menuLabels()).toEqual(["Triple"]);

    clickText("Triple");
    expect(documentXml(handle)).toContain(
      '<w:spacing w:line="720" w:lineRule="auto"/>'
    );
    unmount();
  });

  it("offers the cell borders handed in instead of the built-in three", () => {
    const { handle, unmount } = mount(WITH_TABLE, {
      presets: { cellBorders: [{ label: "Outline only", preset: "outer" }] },
    });
    caretInFirstCell(handle);
    click("Cell borders");

    expect(menuLabels()).toEqual(["Outline only"]);

    clickText("Outline only");
    expect(documentXml(handle)).toContain('<w:top w:val="single" w:sz="4"');
    unmount();
  });

  it("leaves every picker on its built-in list when no presets are handed in", () => {
    const plain = mount(PARAGRAPH);
    selectFirstParagraph(plain.handle);

    expect(optionLabels("Font")).toContain("Trebuchet MS");
    expect(optionLabels("Font size")).toContain("36pt");
    click("Text color");
    expect(swatchLabels()).toHaveLength(80);
    plain.unmount();

    const spaced = mount(PARAGRAPH);
    selectFirstParagraph(spaced.handle);
    click("Line and paragraph spacing");
    expect(menuLabels()).toEqual(["Single", "1.15", "1.5", "Double"]);
    spaced.unmount();

    const tabled = mount(WITH_TABLE);
    caretInFirstCell(tabled.handle);
    click("Cell borders");
    expect(menuLabels()).toEqual([
      "All borders",
      "Outer borders",
      "No borders",
    ]);
    tabled.unmount();
  });
});

const LOCKED_PR =
  '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

const runXml = (text: string) =>
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** Plain text, a locked control over two characters, then plain text again */
const WITH_LOCK = makeDocx(
  `<w:p>${runXml("a")}<w:sdt>${LOCKED_PR}` +
    `<w:sdtContent>${runXml("bc")}</w:sdtContent></w:sdt>` +
    `${runXml("d")}</w:p>`
);

/**
 * A control offered where the guard would refuse the edit is a click that does nothing, so what
 * writes character formatting is drawn from the lock as well
 */
describe("the formatting controls over locked text", () => {
  function selectRange(
    handle: DocxEditorHandle,
    from: number,
    to = from
  ): void {
    act(() => {
      const { view } = handle;
      view.dispatch(
        view.state.tr.setSelection(
          TextSelection.create(view.state.doc, from, to)
        )
      );
    });
  }

  it("go unclickable where the whole selection is shut", () => {
    const { handle, unmount } = mount(WITH_LOCK);
    selectRange(handle, 2, 4);

    expect(button("Bold").disabled).toBe(true);
    expect(button("Italic").disabled).toBe(true);
    expect(button("Text color").disabled).toBe(true);
    expect(button("Highlight").disabled).toBe(true);
    expect(selectBox("Font").disabled).toBe(true);
    expect(selectBox("Font size").disabled).toBe(true);
    // A paragraph holding a lock still takes its own formatting, so those stay clickable
    expect(button("Alignment").disabled).toBe(false);
    expect(button("Increase indent").disabled).toBe(false);
    unmount();
  });

  // Formatting at a caret is staged for the text typed next, which inside a lock cannot go in
  it("go unclickable at a caret inside the shut stretch", () => {
    const { handle, unmount } = mount(WITH_LOCK);
    selectRange(handle, 3);

    expect(button("Bold").disabled).toBe(true);
    selectRange(handle, 2);
    // Against the edge the text typed next lands outside the control
    expect(button("Bold").disabled).toBe(false);
    unmount();
  });

  it("stay clickable where the selection also holds text the lock leaves open", () => {
    const { handle, unmount } = mount(WITH_LOCK);
    selectRange(handle, 1, 5);

    expect(button("Bold").disabled).toBe(false);
    click("Bold");
    expect(documentXml(handle)).toContain("<w:b/>");
    unmount();
  });
});

/**
 * The lock guard refuses the step a replayed lock is made of, so these two buttons have to run the
 * editor's own undo and redo, which carry the pass for it (`editor/commands/historyCommands`)
 */
describe("the history buttons over a lock", () => {
  const LOCK_XML = '<w:lock w:val="sdtContentLocked"/>';

  it("take a lock back and shut it again", () => {
    const { handle, unmount } = mount(PARAGRAPH);
    selectFirstParagraph(handle);
    act(() => {
      lockSelection(handle.view.state, (tr) => handle.view.dispatch(tr));
    });
    expect(documentXml(handle)).toContain(LOCK_XML);

    click("Undo");
    expect(documentXml(handle)).not.toContain(LOCK_XML);

    click("Redo");
    expect(documentXml(handle)).toContain(LOCK_XML);
    unmount();
  });
});

/** A two by two table whose top left cell stands inside a control that shuts the whole cell */
const LOCKED_CELL = makeNumberedDocx(
  '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>' +
    "<w:tbl>" +
    '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>' +
    `<w:tr><w:sdt>${LOCKED_PR}<w:sdtContent>${cellXml("TopLeft")}` +
    `</w:sdtContent></w:sdt>${cellXml("TopRight")}</w:tr>` +
    `<w:tr>${cellXml("BottomLeft")}${cellXml("BottomRight")}</w:tr>` +
    "</w:tbl>"
);

/**
 * Every paragraph edit leaves a locked cell's paragraphs out (`editor/paragraphEdits`), so the
 * commands report that they have nothing to do and the controls that run them are drawn dead
 * without any of them knowing what a lock is
 */
describe("the paragraph controls over a locked cell", () => {
  const PARAGRAPH_CONTROLS = [
    "Alignment",
    "Line and paragraph spacing",
    "Decrease indent",
    "Increase indent",
    "Numbered list",
    "Bulleted list",
  ];

  function caretInCell(handle: DocxEditorHandle, needle: string): void {
    act(() => {
      const { view } = handle;
      let at = -1;
      view.state.doc.descendants((node, pos) => {
        if (at < 0 && node.isText && node.text === needle) at = pos + 1;
      });
      if (at < 0) throw new Error(`text not found: ${needle}`);
      view.dispatch(
        view.state.tr.setSelection(TextSelection.create(view.state.doc, at, at))
      );
    });
  }

  it("go unclickable with the caret inside the cell", () => {
    const { handle, unmount } = mount(LOCKED_CELL);
    caretInCell(handle, "TopLeft");

    for (const label of PARAGRAPH_CONTROLS) {
      expect(button(label).disabled, label).toBe(true);
    }
    unmount();
  });

  it("stay clickable in the cell beside it", () => {
    const { handle, unmount } = mount(LOCKED_CELL);
    caretInCell(handle, "TopRight");

    expect(button("Alignment").disabled).toBe(false);
    expect(button("Line and paragraph spacing").disabled).toBe(false);
    expect(button("Increase indent").disabled).toBe(false);
    expect(button("Numbered list").disabled).toBe(false);
    expect(button("Bulleted list").disabled).toBe(false);
    // A paragraph sitting at the left margin has no indent to give up, lock or no lock
    expect(button("Decrease indent").disabled).toBe(true);
    unmount();
  });

  it("write the edit they offer where the lock does not reach", () => {
    const { handle, unmount } = mount(LOCKED_CELL);
    caretInCell(handle, "TopRight");
    click("Increase indent");

    expect(documentXml(handle)).toContain('<w:ind w:left="720"/>');
    unmount();
  });
});
