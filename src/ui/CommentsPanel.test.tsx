// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import { TextSelection } from "prosemirror-state";
import { act, type ReactNode } from "react";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { decode, makeDocx } from "../__testing__/docx";
import { renderInto } from "../__testing__/react";
import {
  DocxEditor,
  type DocxEditorHandle,
  type DocxEditorMode,
} from "../DocxEditor";
import type { CommentAuthor } from "../editor/commands/commentCommands";

declare global {
  var IS_REACT_ACT_ENVIRONMENT: boolean;
}

const encoder = new TextEncoder();
const REL_BASE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function commentDocument(): Uint8Array {
  const parts = unzipSync(
    makeDocx(
      '<w:p><w:commentRangeStart w:id="2"/>' +
        '<w:r><w:t xml:space="preserve">source</w:t></w:r>' +
        '<w:commentRangeEnd w:id="2"/>' +
        '<w:r><w:commentReference w:id="2"/></w:r></w:p>'
    )
  );
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId2" Target="comments.xml" Type="${REL_BASE}/comments"/>` +
      "</Relationships>"
  );
  parts["word/comments.xml"] = encoder.encode(
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:comment w:id="2" w:author="Ada"><w:p><w:r>' +
      '<w:t xml:space="preserve">Original note</w:t>' +
      "</w:r></w:p></w:comment></w:comments>"
  );
  parts["[Content_Types].xml"] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

function commentReadyDocument(): Uint8Array {
  const parts = unzipSync(
    makeDocx('<w:p><w:r><w:t xml:space="preserve">source</w:t></w:r></w:p>')
  );
  parts["[Content_Types].xml"] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

function datedCommentsDocument(): Uint8Array {
  const parts = unzipSync(
    makeDocx(
      '<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>older</w:t></w:r>' +
        '<w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>' +
        '<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>newer</w:t></w:r>' +
        '<w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r></w:p>'
    )
  );
  parts["word/_rels/document.xml.rels"] = encoder.encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId2" Target="comments.xml" Type="${REL_BASE}/comments"/>` +
      "</Relationships>"
  );
  parts["word/comments.xml"] = encoder.encode(
    '<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:comment w:id="2" w:author="Ada" w:date="2026-08-20T10:00:00Z"><w:p><w:r><w:t>Older note</w:t></w:r></w:p></w:comment>' +
      '<w:comment w:id="3" w:author="Grace" w:date="2026-08-21T10:00:00Z"><w:p><w:r><w:t>Newer note</w:t></w:r></w:p></w:comment>' +
      "</w:comments>"
  );
  parts["[Content_Types].xml"] = encoder.encode(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Override PartName="/word/document.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/comments.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
      "</Types>"
  );
  return zipSync(parts);
}

let host: HTMLDivElement;

beforeEach(() => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true;
  host = document.createElement("div");
  document.body.appendChild(host);
});

afterEach(() => host.remove());

const render = (element: ReactNode) => renderInto(host, element);

/** The identity the tests write under */
const GRACE: CommentAuthor = { id: "grace", name: "Grace", initials: "GH" };

/** A second identity, for whose comments Grace is not the author */
const ADA: CommentAuthor = { id: "ada", name: "Ada" };

const EDITING: DocxEditorMode = { kind: "edit", author: GRACE };

function mount(bytes: Uint8Array, mode: DocxEditorMode = EDITING) {
  const box: { current: DocxEditorHandle | null } = { current: null };
  const unmount = render(
    <DocxEditor
      document={bytes}
      ref={box}
      mode={mode}
      renderImportError={() => null}
    />
  );
  if (!box.current) throw new Error("the editor ref was not attached");
  box.current.view.posAtCoords = () => ({
    pos: box.current?.view.state.selection.head ?? 1,
    inside: -1,
  });
  box.current.view.coordsAtPos = (pos) => ({
    left: 0,
    right: 0,
    top: pos * 10,
    bottom: pos * 10 + 16,
  });
  return { handle: box.current, unmount };
}

function button(name: string): HTMLButtonElement {
  const found = Array.from(host.querySelectorAll("button")).find(
    (element) =>
      element.getAttribute("aria-label") === name ||
      element.textContent === name
  );
  if (!(found instanceof HTMLButtonElement)) {
    throw new Error(`button not found: ${name}`);
  }
  return found;
}

function textarea(name: string): HTMLTextAreaElement {
  const found = host.querySelector(`textarea[aria-label="${name}"]`);
  if (!(found instanceof HTMLTextAreaElement)) {
    throw new Error(`textarea not found: ${name}`);
  }
  return found;
}

function type(field: HTMLTextAreaElement, value: string): void {
  const write = Object.getOwnPropertyDescriptor(
    HTMLTextAreaElement.prototype,
    "value"
  )?.set;
  if (!write) throw new Error("no value setter on HTMLTextAreaElement");
  act(() => {
    write.call(field, value);
    field.dispatchEvent(new Event("input", { bubbles: true }));
  });
}

function click(target: HTMLElement): void {
  act(() => target.click());
}

function selectText(handle: DocxEditorHandle, text: string): void {
  let from = -1;
  handle.view.state.doc.descendants((node, pos) => {
    if (from < 0 && node.isText && node.text?.includes(text)) from = pos;
  });
  if (from < 0) throw new Error(`text not found: ${text}`);
  act(() => {
    const { view } = handle;
    view.dispatch(
      view.state.tr.setSelection(
        TextSelection.create(view.state.doc, from, from + text.length)
      )
    );
  });
}

function rightClickText(): void {
  const paragraph = host.querySelector("p");
  if (!paragraph) throw new Error("paragraph not found");
  act(() => {
    paragraph.dispatchEvent(
      new MouseEvent("contextmenu", {
        bubbles: true,
        cancelable: true,
        clientX: 40,
        clientY: 60,
      })
    );
  });
}

function exportedPart(handle: DocxEditorHandle, path: string): string {
  return decode(unzipSync(handle.exportBytes())[path]);
}

describe("the comments panel", () => {
  it("shows unresolved comments beside the document without opening all comments", () => {
    const { unmount } = mount(commentDocument());

    const panel = host.querySelector('aside[aria-label="Comments"]');
    expect(panel?.getAttribute("data-view")).toBe("rail");
    expect(panel?.textContent).toContain("Original note");
    click(button("Reply to comment: Original note"));
    expect(textarea("Reply text")).toBeTruthy();
    click(button("Cancel"));
    click(button("Resolve"));
    expect(host.querySelector('aside[aria-label="Comments"]')).toBeNull();
    unmount();
  });

  it("lists all comments from newest to oldest", () => {
    const { unmount } = mount(datedCommentsDocument());
    click(button("Show comments"));

    const bodies = Array.from(
      host.querySelectorAll(`.${"docx-editor-comment-body"}`)
    ).map((element) => element.textContent);
    expect(bodies).toEqual(["Newer note", "Older note"]);
    unmount();
  });

  it("keeps the all-comments scroll position across selection changes", () => {
    const { unmount } = mount(commentDocument());
    click(button("Show comments"));
    const panel = host.querySelector('aside[aria-label="Comments"]');
    const meta = host.querySelector(`button.${"docx-editor-comment-meta"}`);
    if (
      !(panel instanceof HTMLElement) ||
      !(meta instanceof HTMLButtonElement)
    ) {
      throw new Error("the comments panel was not drawn");
    }
    panel.scrollTop = 80;
    click(meta);

    expect(panel.scrollTop).toBe(80);
    unmount();
  });

  it("lets a read-only editor open resolved comment history", () => {
    const editable = mount(commentDocument());
    click(button("Show comments"));
    click(button("Resolve"));
    const resolved = editable.handle.exportBytes();
    editable.unmount();

    const readOnly = mount(resolved, { kind: "readOnly" });
    click(button("Show comments"));
    expect(
      host.querySelector('aside[aria-label="Comments"]')?.textContent
    ).toContain("Original note");
    readOnly.unmount();
  });

  it("shows imported comments and writes an edited body", () => {
    const { handle, unmount } = mount(commentDocument());
    click(button("Show comments"));

    expect(
      host.querySelector('aside[aria-label="Comments"]')?.textContent
    ).toContain("Original note");
    click(button("Edit"));
    type(textarea("Edit comment"), "Revised note");
    click(button("Save"));

    expect(exportedPart(handle, "word/comments.xml")).toContain("Revised note");
    unmount();
  });

  it("closes an unchanged edit without rewriting the comment", () => {
    const { handle, unmount } = mount(commentDocument());
    click(button("Show comments"));
    const before = exportedPart(handle, "word/comments.xml");

    click(button("Edit"));
    click(button("Save"));

    expect(
      host.querySelector('textarea[aria-label="Edit comment"]')
    ).toBeNull();
    expect(exportedPart(handle, "word/comments.xml")).toBe(before);
    unmount();
  });

  it("deletes the body and every anchor marker", () => {
    const { handle, unmount } = mount(commentDocument());
    click(button("Show comments"));
    click(button("Delete"));

    expect(host.textContent).toContain("This document has no comments.");
    expect(exportedPart(handle, "word/document.xml")).not.toContain(
      "commentReference"
    );
    expect(exportedPart(handle, "word/comments.xml")).not.toContain(
      "w:comment "
    );
    unmount();
  });

  it("opens the composer only for a text selection and adds the configured author", () => {
    const { handle, unmount } = mount(commentReadyDocument());
    expect(button("Show comments").disabled).toBe(false);
    selectText(handle, "source");
    rightClickText();
    expect(button("Add comment").disabled).toBe(false);
    click(button("Add comment"));
    type(textarea("Comment text"), "New note");
    click(button("Comment"));

    expect(host.textContent).toContain("New note");
    expect(exportedPart(handle, "word/comments.xml")).toContain(
      'w:author="Grace"'
    );
    unmount();
  });

  it("adds a reply and resolves or reopens its thread", () => {
    const { handle, unmount } = mount(commentDocument());
    click(button("Show comments"));
    click(button("Reply to comment: Original note"));
    type(textarea("Reply text"), "Follow-up");
    click(button("Reply"));

    expect(host.textContent).toContain("Follow-up");
    click(button("Resolve"));
    expect(host.textContent).toContain("Original note");
    expect(host.querySelector('button[aria-label="Edit"]')).toBeNull();
    expect(host.querySelector('button[aria-label="Delete"]')).toBeNull();
    expect(
      host.querySelector('button[aria-label^="Reply to comment:"]')
    ).toBeNull();
    click(button("Reopen"));
    expect(host.textContent).toContain("Original note");
    expect(exportedPart(handle, "word/commentsExtended.xml")).toContain(
      'w15:done="0"'
    );
    unmount();
  });
});

describe("whose comments the panel offers to edit", () => {
  /** A document holding one comment Ada wrote, made the way the panel makes one */
  function adaCommented(): Uint8Array {
    const { handle, unmount } = mount(commentReadyDocument(), {
      kind: "edit",
      author: ADA,
    });
    selectText(handle, "source");
    rightClickText();
    click(button("Add comment"));
    type(textarea("Comment text"), "Ada's note");
    click(button("Comment"));
    const bytes = handle.exportBytes();
    unmount();
    return bytes;
  }

  const editButtons = () =>
    Array.from(host.querySelectorAll("button[aria-label]"))
      .map((element) => element.getAttribute("aria-label"))
      .filter((label) => label === "Edit" || label === "Delete");

  it("records the identity the comment was written under", () => {
    const { handle, unmount } = mount(adaCommented(), {
      kind: "edit",
      author: ADA,
    });
    expect(exportedPart(handle, "word/people.xml")).toContain(
      'w15:userId="ada"'
    );
    unmount();
  });

  it("offers one's own comment for editing and deleting", () => {
    const { unmount } = mount(adaCommented(), { kind: "edit", author: ADA });
    click(button("Show comments"));
    expect(editButtons()).toEqual(["Edit", "Delete"]);
    unmount();
  });

  it("offers another author's comment for replying and settling, not for editing", () => {
    const { handle, unmount } = mount(adaCommented());
    click(button("Show comments"));
    expect(editButtons()).toEqual([]);

    click(button("Reply to comment: Ada's note"));
    type(textarea("Reply text"), "Seen, thanks");
    click(button("Reply"));
    expect(host.textContent).toContain("Seen, thanks");
    // The reply is Grace's own, so it is hers to edit
    expect(
      host.querySelector('button[aria-label="Edit reply"]')
    ).not.toBeNull();

    click(button("Resolve"));
    expect(exportedPart(handle, "word/commentsExtended.xml")).toContain(
      'w15:done="1"'
    );
    unmount();
  });

  it("opens every comment to a moderator", () => {
    const { unmount } = mount(adaCommented(), {
      kind: "edit",
      author: GRACE,
      editableComments: "all",
    });
    click(button("Show comments"));
    expect(editButtons()).toEqual(["Edit", "Delete"]);
    unmount();
  });

  /** `commentDocument` holds a comment by an Ada no people part vouches for, which is everyone's */
  it("leaves a comment carrying no identity open to everyone", () => {
    const { unmount } = mount(commentDocument());
    click(button("Show comments"));
    expect(editButtons()).toEqual(["Edit", "Delete"]);
    unmount();
  });
});

describe("a commenter", () => {
  const COMMENTING: DocxEditorMode = { kind: "comment", author: GRACE };

  it("writes a comment on the selected text under its own identity, and changes nothing else", () => {
    const { handle, unmount } = mount(commentReadyDocument(), COMMENTING);
    selectText(handle, "source");
    rightClickText();
    click(button("Add comment"));
    type(textarea("Comment text"), "Please check");
    click(button("Comment"));

    expect(host.textContent).toContain("Please check");
    expect(handle.view.state.doc.textContent).toBe("source");
    expect(exportedPart(handle, "word/comments.xml")).toContain(
      'w:author="Grace"'
    );
    expect(exportedPart(handle, "word/people.xml")).toContain(
      'w15:userId="grace"'
    );
    unmount();
  });

  it("is offered no composer as a reader", () => {
    const { handle, unmount } = mount(commentReadyDocument(), {
      kind: "readOnly",
    });
    selectText(handle, "source");
    rightClickText();
    expect(host.querySelector('button[role="menuitem"]')).toBeNull();
    unmount();
  });
});
