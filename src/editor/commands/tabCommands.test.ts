// @vitest-environment jsdom
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { makeDocx } from "../../__testing__/docx";
import { runCommand, select } from "../../__testing__/editing";
import { importDocx } from "../../docx/importDocx";
import { serializeParagraph } from "../../docx/serializeParagraph";
import { docxSchema } from "../../schema";
import { createEditorState } from "../createEditor";
import { insertTab, moveAcrossTab } from "./tabCommands";

function opened(): EditorState {
  const { doc } = importDocx(
    makeDocx('<w:p><w:r><w:t xml:space="preserve">abcd</w:t></w:r></w:p>')
  );
  return createEditorState(doc);
}

describe("insertTab", () => {
  it("inserts one text position that serializes as w:tab", () => {
    const next = runCommand(select(opened(), 3), insertTab);
    const paragraph = next.doc.child(0);

    expect(paragraph.textContent).toBe("ab\tcd");
    let tabCount = 0;
    paragraph.descendants((node) => {
      if (
        node.isText &&
        node.text === "\t" &&
        node.marks.some((mark) => mark.type.name === "tab")
      ) {
        tabCount += 1;
      }
    });
    expect(tabCount).toBe(1);
    expect(serializeParagraph(paragraph)).toContain("<w:tab/>");
  });

  it("replaces a selection inside one paragraph", () => {
    const next = runCommand(select(opened(), 2, 4), insertTab);
    expect(next.doc.child(0).textContent).toBe("a\td");
  });

  it("does not replace a selection spanning paragraphs", () => {
    const state = createEditorState(
      importDocx(
        makeDocx(
          "<w:p><w:r><w:t>a</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>b</w:t></w:r></w:p>"
        )
      ).doc
    );
    expect(insertTab(select(state, 1, state.doc.content.size - 1))).toBe(false);
  });

  it("leaves list paragraphs to their level-changing controls", () => {
    const paragraph = docxSchema.nodes.paragraph.create(
      { format: { numbering: { numId: 1, ilvl: 0 } } },
      docxSchema.text("item")
    );
    const state = select(
      createEditorState(docxSchema.nodes.doc.create(null, paragraph)),
      2
    );

    expect(insertTab(state)).toBe(false);
  });

  it("keeps a link around a tab inserted inside its text", () => {
    const link = docxSchema.marks.link.create({
      href: "https://example.com",
      linkKey: 1,
    });
    const doc = docxSchema.nodes.doc.create(null, [
      docxSchema.nodes.paragraph.create(null, docxSchema.text("ab", [link])),
    ]);
    const next = runCommand(select(createEditorState(doc), 2), insertTab);
    const tab = next.doc.child(0).child(1);

    expect(tab.text).toBe("\t");
    expect(tab.marks.some((mark) => mark.type === docxSchema.marks.link)).toBe(
      true
    );
  });
});

describe("moveAcrossTab", () => {
  const tabbed = () => runCommand(select(opened(), 3), insertTab);

  it("moves and extends by exactly one tab character", () => {
    const afterTab = select(tabbed(), 4);
    const before = runCommand(afterTab, moveAcrossTab("left", false));
    expect(before.selection.from).toBe(3);

    const selected = runCommand(before, moveAcrossTab("right", true));
    expect(selected.selection.from).toBe(3);
    expect(selected.selection.to).toBe(4);
  });

  it("leaves ordinary characters to the browser", () => {
    expect(moveAcrossTab("left", false)(select(opened(), 3))).toBe(false);
  });
});
