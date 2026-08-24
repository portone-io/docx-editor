// @vitest-environment jsdom
import type { Node as PMNode } from "prosemirror-model";
import { type EditorState, TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { makeDocx, TINY_PNG_DATA_URL } from "../__testing__/docx";
import { posOfText } from "../__testing__/editing";
import { importDocx } from "../docx/importDocx";
import { createEditorState } from "./createEditor";
import { canInsertImage, type ImageToInsert, insertImage } from "./insertImage";

const BODY = '<w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p>';

const BOLD_BODY =
  "<w:p><w:r><w:rPr><w:b/></w:rPr>" +
  '<w:t xml:space="preserve">Body</w:t></w:r></w:p>';

const cellXml = (text: string) =>
  `<w:tc><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;

const TABLE_BODY =
  "<w:tbl>" +
  '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cellXml("Left")}${cellXml("Right")}</w:tr>` +
  "</w:tbl>";

/** A body holding nothing the editor can model, so there is no inline position anywhere */
const ONLY_PRESERVED = "<w:sdt><w:sdtContent/></w:sdt>";

const A_PICTURE: ImageToInsert = {
  src: TINY_PNG_DATA_URL,
  extent: { cx: 1905000, cy: 952500 },
  alt: "a picture",
};

/** An editing state with the caret sitting just before a piece of text */
function openWithCaret(body: string, needle: string): EditorState {
  const { doc } = importDocx(makeDocx(body));
  const state = createEditorState(doc);
  return state.apply(
    state.tr.setSelection(TextSelection.create(doc, posOfText(doc, needle)))
  );
}

/** Runs the command and returns the next state. Null when there was nothing to do */
function run(state: EditorState, image: ImageToInsert): EditorState | null {
  let next: EditorState | null = null;
  const handled = insertImage(image)(state, (tr) => {
    next = state.apply(tr);
  });
  return handled ? next : null;
}

function insertedImage(state: EditorState, image: ImageToInsert): PMNode {
  const next = run(state, image);
  if (!next) throw new Error("could not insert the image");
  const found: PMNode[] = [];
  next.doc.descendants((node) => {
    if (node.type.name === "image") found.push(node);
    return true;
  });
  if (found.length !== 1) throw new Error(`images found: ${found.length}`);
  return found[0];
}

describe("the insert image command", () => {
  it("puts the image at the caret, carrying the bytes and the size it was handed", () => {
    const state = openWithCaret(BODY, "Body");
    expect(canInsertImage(state)).toBe(true);

    const image = insertedImage(state, A_PICTURE);
    expect(image.attrs.src).toBe(TINY_PNG_DATA_URL);
    expect(image.attrs.extent).toEqual({ cx: 1905000, cy: 952500 });
    expect(image.attrs.alt).toBe("a picture");
  });

  it("leaves the image without original XML, which is what marks it as inserted", () => {
    const state = openWithCaret(BODY, "Body");
    expect(insertedImage(state, A_PICTURE).attrs.xml).toBe(null);
  });

  it("takes no alternative text as none at all", () => {
    const state = openWithCaret(BODY, "Body");
    const image = insertedImage(state, {
      src: A_PICTURE.src,
      extent: A_PICTURE.extent,
    });
    expect(image.attrs.alt).toBe(null);
  });

  it("keeps the text around the caret and leaves the caret past the image", () => {
    // The caret sits between the B and the ody, so the image goes in between them
    const state = openWithCaret(BODY, "Body");
    const next = run(state, A_PICTURE);
    if (!next) throw new Error("could not insert the image");

    expect(next.doc.childCount).toBe(1);
    expect(next.doc.textContent).toBe("Body");
    const kinds = next.doc.child(0).children.map((child) => child.type.name);
    expect(kinds).toEqual(["text", "image", "text"]);
    // Right after the image, so typing carries on where the picture was put
    expect(next.selection.from).toBe(3);
  });

  it("goes in wearing the formatting of the run it landed in", () => {
    const state = openWithCaret(BOLD_BODY, "Body");
    const image = insertedImage(state, A_PICTURE);
    const mark = image.marks.find((entry) => entry.type.name === "run");

    expect(mark?.attrs.rPr).toBe("<w:rPr><w:b/></w:rPr>");
  });

  it("replaces the selected text", () => {
    const opened = openWithCaret(BODY, "Body");
    const selected = opened.apply(
      opened.tr.setSelection(TextSelection.create(opened.doc, 1, 5))
    );

    const next = run(selected, A_PICTURE);
    if (!next) throw new Error("could not insert the image");
    expect(next.doc.textContent).toBe("");
    expect(next.doc.child(0).child(0).type.name).toBe("image");
  });

  it("goes into a table cell as it does into the body", () => {
    const state = openWithCaret(TABLE_BODY, "Left");
    expect(canInsertImage(state)).toBe(true);
    expect(insertedImage(state, A_PICTURE).attrs.src).toBe(TINY_PNG_DATA_URL);
  });

  it("refuses a src that is not the bytes of an image we can draw", () => {
    const state = openWithCaret(BODY, "Body");
    for (const src of [
      "https://example.com/picture.png",
      "javascript:alert(1)",
      "data:image/tiff;base64,AAAA",
      "data:image/png;base64,not base64",
      "",
    ]) {
      expect(insertImage({ ...A_PICTURE, src })(state)).toBe(false);
    }
  });

  it("refuses a size that cannot be shown", () => {
    const state = openWithCaret(BODY, "Body");
    for (const extent of [
      { cx: 0, cy: 100 },
      { cx: 100, cy: 0 },
      { cx: -100, cy: 100 },
      { cx: 1.5, cy: 100 },
    ]) {
      expect(insertImage({ ...A_PICTURE, extent })(state)).toBe(false);
    }
  });

  it("has nowhere to go in a document made up entirely of preserved blocks", () => {
    const { doc } = importDocx(makeDocx(ONLY_PRESERVED));
    const state = createEditorState(doc);

    expect(state.doc.child(0).type.name).toBe("docxRaw");
    expect(canInsertImage(state)).toBe(false);
    expect(insertImage(A_PICTURE)(state)).toBe(false);
  });
});
