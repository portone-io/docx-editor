// @vitest-environment jsdom
import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import {
  decode,
  documentXmlOf,
  makeDocx,
  makeLinkedDocx,
} from "../../__testing__/docx";
import { runCommand, select } from "../../__testing__/editing";
import { exportDocx } from "../../docx/exportDocx";
import { importDocx } from "../../docx/importDocx";
import type { SessionStore } from "../../docx/session";
import { docxSchema } from "../../schema";
import { createEditorState } from "../createEditor";
import {
  activeLink,
  activeLinkSpan,
  canSetLink,
  removeLink,
  setLink,
} from "./linkCommands";
import { lockSelection } from "./lockCommands";

const run = (text: string, rPr = "") =>
  `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;

const TERMS = "https://example.com/terms";
const PRICES = "https://example.com/prices";

interface Opened {
  state: EditorState;
  session: SessionStore;
}

function opened(bytes: Uint8Array): Opened {
  const { doc, session } = importDocx(bytes);
  return { state: createEditorState(doc), session };
}

/** A document with the one paragraph and no relationships at all */
function plain(body: string): Opened {
  return opened(makeDocx(body));
}

/** The same, with `rId9` relating an address the body's links can point at */
function linked(body: string, target = TERMS): Opened {
  return opened(makeLinkedDocx(body, { rId9: target }));
}

/** The stretch the first occurrence of this text covers */
function rangeOf(doc: PMNode, needle: string): { from: number; to: number } {
  let found = -1;
  doc.descendants((node, pos) => {
    if (found >= 0 || !node.isText) return true;
    const at = node.text?.indexOf(needle) ?? -1;
    if (at >= 0) found = pos + at;
    return true;
  });
  if (found < 0) throw new Error(`text not found: ${needle}`);
  return { from: found, to: found + needle.length };
}

/** The whole of that text selected */
function over(state: EditorState, needle: string): EditorState {
  const { from, to } = rangeOf(state.doc, needle);
  return select(state, from, to);
}

/** The caret one character into that text */
function inside(state: EditorState, needle: string): EditorState {
  return select(state, rangeOf(state.doc, needle).from + 1);
}

/** Whether anything in the document still wears a link mark */
function anyLinkMark(state: EditorState): boolean {
  let found = false;
  state.doc.descendants((node) => {
    if (node.marks.some((mark) => mark.type === docxSchema.marks.link)) {
      found = true;
    }
    return !found;
  });
  return found;
}

function bodyXml(state: EditorState, session: SessionStore): string {
  return documentXmlOf(state.doc, session);
}

/** Every hyperlink opening tag the export wrote */
function linkTags(state: EditorState, session: SessionStore): string[] {
  return bodyXml(state, session).match(/<w:hyperlink[^>]*>/g) ?? [];
}

/** The relationships part as the export wrote it */
function relsXml(state: EditorState, session: SessionStore): string {
  const out = unzipSync(exportDocx(state.doc, session));
  return decode(out["word/_rels/document.xml.rels"]);
}

const RELATIONSHIP = /<Relationship[^>]*\/>/g;

function relationships(state: EditorState, session: SessionStore): string[] {
  return relsXml(state, session).match(RELATIONSHIP) ?? [];
}

/** The `Relationship` entries the export added on top of what the document came with */
function addedRels(before: string, after: string): string[] {
  const had: string[] = before.match(RELATIONSHIP) ?? [];
  return (after.match(RELATIONSHIP) ?? []).filter(
    (entry) => !had.includes(entry)
  );
}

describe("putting a link on plain text", () => {
  const BODY = `<w:p>${run("see our terms")}</w:p>`;

  it("wraps the selected text in a hyperlink pointing at a new relationship", () => {
    const { state, session } = plain(BODY);
    const linkedState = runCommand(over(state, "our terms"), setLink(TERMS));
    const xml = bodyXml(linkedState, session);

    expect(xml).toContain(`${run("see ")}<w:hyperlink`);
    expect(xml).toMatch(
      /<w:hyperlink xmlns:r="[^"]+" r:id="rId\d+"><w:r><w:t xml:space="preserve">our terms<\/w:t><\/w:r><\/w:hyperlink>/
    );
  });

  it("writes exactly one relationship for it, marked as pointing outside the package", () => {
    const { state, session } = plain(BODY);
    const linkedState = runCommand(over(state, "our terms"), setLink(TERMS));
    const rels = relsXml(linkedState, session);

    expect(relationships(linkedState, session)).toHaveLength(1);
    expect(rels).toContain(`Target="${TERMS}"`);
    expect(rels).toContain('TargetMode="External"');
    expect(rels).toContain("/hyperlink");
  });

  it("escapes an address that carries markup characters", () => {
    const { state, session } = plain(BODY);
    const address = "https://example.com/s?a=1&b=2";
    const linkedState = runCommand(over(state, "our terms"), setLink(address));

    expect(relsXml(linkedState, session)).toContain(
      'Target="https://example.com/s?a=1&amp;b=2"'
    );
  });

  it("points a second link at the same address at the one relationship", () => {
    const { state, session } = plain(`<w:p>${run("a and b")}</w:p>`);
    const first = runCommand(over(state, "a"), setLink(TERMS));
    const second = runCommand(over(first, "b"), setLink(TERMS));

    expect(relationships(second, session)).toHaveLength(1);
    expect(linkTags(second, session)).toHaveLength(2);
  });

  it("leaves the run's own formatting where it was", () => {
    const { state, session } = plain(
      `<w:p>${run("terms", "<w:rPr><w:b/></w:rPr>")}</w:p>`
    );
    const linkedState = runCommand(over(state, "terms"), setLink(TERMS));

    expect(bodyXml(linkedState, session)).toContain(
      `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">terms</w:t></w:r>`
    );
  });

  it("does nothing to a caret standing in no link", () => {
    const { state } = plain(BODY);
    const at = inside(state, "see our terms");
    expect(canSetLink(at)).toBe(false);
    expect(setLink(TERMS)(at)).toBe(false);
  });

  it("refuses an address of nothing but space", () => {
    const { state } = plain(BODY);
    expect(setLink("   ")(over(state, "our terms"))).toBe(false);
  });

  /**
   * A target with no scheme is read as relative to the package, so a bare host would name a file
   * beside the document. The scheme is settled as the link is made, the way Word and Google Docs
   * settle it, rather than papered over wherever the address is opened.
   */
  it.each([
    ["  example.com/terms  ", "http://example.com/terms"],
    ["www.example.com", "http://www.example.com"],
    ["example.com?a=1", "http://example.com?a=1"],
  ])("writes %s as %s", (typed, written) => {
    const { state, session } = plain(BODY);
    const linkedState = runCommand(over(state, "our terms"), setLink(typed));
    expect(relsXml(linkedState, session)).toContain(`Target="${written}"`);
  });

  it.each([
    ["https://example.com/terms", "an address naming its own scheme"],
    ["mailto:terms@example.com", "an address that is not a web address"],
    ["../terms.docx", "a file beside the document"],
    ["#chapter3", "a place in the document"],
    ["intranet", "a name that is no host"],
  ])("writes %s as it reads, being %s", (typed) => {
    const { state, session } = plain(BODY);
    const linkedState = runCommand(over(state, "our terms"), setLink(typed));
    expect(relsXml(linkedState, session)).toContain(`Target="${typed}"`);
  });
});

describe("the address at the selection", () => {
  const BODY =
    "<w:p>" +
    run("see ") +
    `<w:hyperlink r:id="rId9">${run("our terms")}</w:hyperlink>` +
    run(" and ") +
    `<w:hyperlink w:anchor="chapter3">${run("chapter three")}</w:hyperlink>` +
    "</w:p>";

  it("is the address of the link the caret stands in", () => {
    const { state } = linked(BODY);
    expect(activeLink(inside(state, "our terms"))).toBe(TERMS);
  });

  it("is nothing where the caret stands in no link", () => {
    const { state } = linked(BODY);
    expect(activeLink(inside(state, "see "))).toBeNull();
  });

  it("is nothing at the very edge of a link, where a typed character lands outside it", () => {
    const { state } = linked(BODY);
    const { from, to } = rangeOf(state.doc, "our terms");
    expect(activeLink(select(state, from))).toBeNull();
    expect(activeLink(select(state, to))).toBeNull();
  });

  it("is nothing for a link that names a bookmark, since it has no address", () => {
    const { state } = linked(BODY);
    expect(activeLink(inside(state, "chapter three"))).toBeNull();
  });

  it("is nothing where the selection runs over several links", () => {
    const { state } = linked(BODY);
    const from = rangeOf(state.doc, "see ").from;
    expect(activeLink(select(state, from, from + 20))).toBeNull();
  });

  it("is the one address a selection covering a single link holds", () => {
    const { state } = linked(BODY);
    expect(activeLink(over(state, "our terms"))).toBe(TERMS);
  });
});

describe("the link the selection sits inside", () => {
  const BODY =
    "<w:p>" +
    run("see ") +
    `<w:hyperlink r:id="rId9">${run("our terms")}</w:hyperlink>` +
    run(" and ") +
    `<w:hyperlink w:anchor="chapter3">${run("chapter three")}</w:hyperlink>` +
    "</w:p>";

  it("is the stretch the link covers and the address it points at", () => {
    const { state } = linked(BODY);
    expect(activeLinkSpan(inside(state, "our terms"))).toEqual({
      ...rangeOf(state.doc, "our terms"),
      href: TERMS,
    });
  });

  it("is the same stretch for a selection that does not leave the link", () => {
    const { state } = linked(BODY);
    const { from, to } = rangeOf(state.doc, "our terms");
    expect(activeLinkSpan(over(state, "our terms"))).toEqual({
      from,
      to,
      href: TERMS,
    });
    expect(activeLinkSpan(select(state, from + 1, to - 1))?.from).toBe(from);
  });

  it("is nothing where the caret stands in no link, and nothing at either edge", () => {
    const { state } = linked(BODY);
    const { from, to } = rangeOf(state.doc, "our terms");
    expect(activeLinkSpan(inside(state, "see "))).toBeNull();
    expect(activeLinkSpan(select(state, from))).toBeNull();
    expect(activeLinkSpan(select(state, to))).toBeNull();
  });

  it("is nothing for a selection that runs past the edge of a link", () => {
    const { state } = linked(BODY);
    const { to } = rangeOf(state.doc, "our terms");
    expect(activeLinkSpan(select(state, to - 2, to + 1))).toBeNull();
  });

  it("is nothing for a selection running over two links", () => {
    const { state } = linked(BODY);
    const from = rangeOf(state.doc, "our terms").from;
    const to = rangeOf(state.doc, "chapter three").to;
    expect(activeLinkSpan(select(state, from + 1, to - 1))).toBeNull();
  });

  /** A link naming a bookmark is a link with nowhere to show, which is not the same as no link */
  it("is the stretch with no address for a link naming a bookmark", () => {
    const { state } = linked(BODY);
    expect(activeLinkSpan(inside(state, "chapter three"))).toEqual({
      ...rangeOf(state.doc, "chapter three"),
      href: null,
    });
  });

  /** What may be done to it is the commands' answer; the address is worth reading either way */
  it("reads a link inside a locked control, which the commands refuse to edit", () => {
    const control =
      '<w:sdt><w:sdtPr><w:id w:val="7"/>' +
      '<w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>' +
      `<w:hyperlink r:id="rId9">${run("settled")}</w:hyperlink>` +
      "</w:sdtContent></w:sdt>";
    const { state } = linked(`<w:p>${control}</w:p>`);
    const shut = inside(state, "settled");

    expect(activeLinkSpan(shut)?.href).toBe(TERMS);
    expect(activeLink(shut)).toBeNull();
    expect(removeLink(shut)).toBe(false);
  });
});

describe("changing where a link points", () => {
  const BODY =
    "<w:p>" +
    '<w:hyperlink r:id="rId9" w:tooltip="Our terms" w:history="1">' +
    `${run("our terms")}</w:hyperlink>` +
    "</w:p>";

  it("writes a new relationship and points the link at it", () => {
    const { state, session } = linked(BODY);
    const before = relsXml(state, session);
    const changed = runCommand(inside(state, "our terms"), setLink(PRICES));
    const added = addedRels(before, relsXml(changed, session));

    expect(added).toHaveLength(1);
    expect(added[0]).toContain(`Target="${PRICES}"`);
    // The relationship the link arrived on is left where it was: additions only
    expect(relsXml(changed, session)).toContain(`Target="${TERMS}"`);
  });

  it("keeps everything else the wrapper carried", () => {
    const { state, session } = linked(BODY);
    const changed = runCommand(inside(state, "our terms"), setLink(PRICES));
    const [tag] = linkTags(changed, session);

    expect(tag).toContain('w:tooltip="Our terms"');
    expect(tag).toContain('w:history="1"');
    expect(tag).not.toContain('r:id="rId9"');
  });

  it("changes the whole link from a caret standing anywhere in it", () => {
    const { state, session } = linked(BODY);
    const changed = runCommand(inside(state, "our terms"), setLink(PRICES));

    expect(linkTags(changed, session)).toHaveLength(1);
    expect(bodyXml(changed, session)).toContain("our terms</w:t>");
  });

  it("reports nothing to do when the address is already that one", () => {
    const { state } = linked(BODY);
    expect(setLink(TERMS)(inside(state, "our terms"))).toBe(false);
  });

  /**
   * A drag selection inside a link retargets the link rather than cutting a second address out of
   * the middle of it, which is what the caret does and what neither reference editor departs from.
   */
  it("changes the whole link where only part of it is selected", () => {
    const { state, session } = linked(BODY);
    const changed = runCommand(over(state, "our"), setLink(PRICES));

    expect(linkTags(changed, session)).toHaveLength(1);
    expect(activeLink(over(changed, "our terms"))).toBe(PRICES);
  });

  it("links the plain text beside a link and changes that link whole", () => {
    const body = `<w:p>${run("see ")}<w:hyperlink r:id="rId9">${run("our terms")}</w:hyperlink></w:p>`;
    const { state, session } = linked(body);
    const { from } = rangeOf(state.doc, "see ");
    const changed = runCommand(select(state, from, from + 7), setLink(PRICES));
    const whole = select(changed, from, from + "see our terms".length);

    expect(activeLink(whole)).toBe(PRICES);
    expect(linkTags(changed, session)).toHaveLength(2);
  });
});

describe("taking a link off", () => {
  const BODY =
    "<w:p>" +
    run("see ") +
    `<w:hyperlink r:id="rId9">${run("our terms", "<w:rPr><w:b/></w:rPr>")}` +
    "</w:hyperlink></w:p>";

  it("leaves the text and its formatting, and drops the wrapper", () => {
    const { state, session } = linked(BODY);
    const bare = runCommand(inside(state, "our terms"), removeLink);
    const xml = bodyXml(bare, session);

    expect(xml).not.toContain("w:hyperlink");
    expect(xml).toContain(
      '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">our terms</w:t></w:r>'
    );
    expect(bare.doc.child(0).textContent).toBe("see our terms");
  });

  /**
   * A link is one object wrapping a stretch rather than formatting spread over it, so however
   * little of it the selection touches, the whole of it comes off. Taking off the selected part
   * alone left the text in three: link, plain, link.
   */
  it("takes the whole link off from a selection covering part of it", () => {
    const body = `<w:p><w:hyperlink r:id="rId9">${run("markup of every")}</w:hyperlink></w:p>`;
    const { state, session } = linked(body);
    const bare = runCommand(over(state, "p of"), removeLink);

    expect(linkTags(bare, session)).toEqual([]);
    expect(anyLinkMark(bare)).toBe(false);
    expect(bare.doc.child(0).textContent).toBe("markup of every");
  });

  it("takes both links off where one selection touches two", () => {
    const body =
      "<w:p>" +
      `<w:hyperlink r:id="rId9">${run("our terms")}</w:hyperlink>` +
      run(" and ") +
      `<w:hyperlink r:id="rId9">${run("our prices")}</w:hyperlink>` +
      "</w:p>";
    const { state, session } = linked(body);
    // One character into the second link, which is where it starts being touched at all
    const from = rangeOf(state.doc, "terms").from;
    const bare = runCommand(select(state, from, from + 11), removeLink);

    expect(linkTags(bare, session)).toEqual([]);
    expect(anyLinkMark(bare)).toBe(false);
  });

  it("reports nothing to do where there is no link", () => {
    const { state } = linked(BODY);
    expect(removeLink(inside(state, "see "))).toBe(false);
    expect(removeLink(over(state, "see "))).toBe(false);
  });

  /** The mark is not inclusive, so a caret at either edge stands outside the link, as `activeLink` reads it */
  it("reports nothing to do for a caret at either edge of a link", () => {
    const { state } = linked(BODY);
    const { from, to } = rangeOf(state.doc, "our terms");
    expect(removeLink(select(state, from))).toBe(false);
    expect(removeLink(select(state, to))).toBe(false);
  });
});

/**
 * A lock over part of a link is the one way a single link comes to run into locked text: a link
 * arriving with a control inside it is preserved whole rather than read (`spec/notes/hyperlinks.md`).
 * Half a link cannot be taken off or retargeted without recreating the split this all came from, so
 * a link a lock holds any part of is left alone entirely.
 */
describe("a link a lock holds part of", () => {
  const BODY =
    "<w:p>" +
    run("see ") +
    `<w:hyperlink r:id="rId9">${run("our terms")}</w:hyperlink>` +
    "</w:p>";

  function halfLocked(): Opened {
    const { state, session } = linked(BODY);
    return { state: runCommand(over(state, "our"), lockSelection), session };
  }

  it("is left whole rather than half taken off", () => {
    const { state } = halfLocked();
    expect(removeLink(inside(state, "terms"))).toBe(false);
    expect(anyLinkMark(state)).toBe(true);
    expect(activeLink(inside(state, "terms"))).toBeNull();
  });

  it("is left pointing where it did rather than half retargeted", () => {
    const { state } = halfLocked();
    expect(setLink(PRICES)(inside(state, "terms"))).toBe(false);
  });

  it("leaves the plain text beside it linkable", () => {
    const { state } = halfLocked();
    expect(canSetLink(over(state, "see "))).toBe(true);
  });
});

describe("typing inside a link", () => {
  const BODY = `<w:p><w:hyperlink r:id="rId9">${run("our terms")}</w:hyperlink></w:p>`;

  it("leaves both halves of a split paragraph linked to the one address", () => {
    const { state, session } = linked(BODY);
    const at = rangeOf(state.doc, "our terms").from + 4;
    const split = state.apply(state.tr.split(at));

    expect(linkTags(split, session)).toEqual([
      '<w:hyperlink r:id="rId9">',
      '<w:hyperlink r:id="rId9">',
    ]);
    expect(relationships(split, session)).toHaveLength(1);
  });

  it("keeps the text typed in the middle inside the one link", () => {
    const { state, session } = linked(BODY);
    const at = rangeOf(state.doc, "our terms").from + 4;
    const typed = state.apply(state.tr.insertText("own ", at));

    expect(typed.doc.child(0).textContent).toBe("our own terms");
    expect(linkTags(typed, session)).toHaveLength(1);
  });
});

describe("a link over locked content", () => {
  const lockedControl = (inner: string) =>
    '<w:sdt><w:sdtPr><w:id w:val="7"/>' +
    '<w:lock w:val="sdtContentLocked"/></w:sdtPr>' +
    `<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

  const BODY = `<w:p>${run("open ")}${lockedControl(run("settled"))}${run(" open")}</w:p>`;

  it("edits the open pieces and leaves the locked stretch out", () => {
    const { state, session } = plain(BODY);
    const whole = select(state, 1, state.doc.child(0).content.size + 1);
    expect(canSetLink(whole)).toBe(true);

    const linkedState = runCommand(whole, setLink(TERMS));
    const xml = bodyXml(linkedState, session);

    expect(xml).toContain("<w:sdtContent><w:r>");
    expect(linkTags(linkedState, session)).toHaveLength(2);
  });

  it("answers false where the lock leaves nothing of the selection", () => {
    const { state } = plain(BODY);
    const { from, to } = rangeOf(state.doc, "settled");
    const shut = select(state, from, to);

    expect(canSetLink(shut)).toBe(false);
    expect(setLink(TERMS)(shut)).toBe(false);
    expect(activeLink(shut)).toBeNull();
  });

  it("takes a link off the open pieces alone", () => {
    const control = lockedControl(
      `<w:hyperlink r:id="rId9">${run("settled")}</w:hyperlink>`
    );
    const { state, session } = linked(
      `<w:p><w:hyperlink r:id="rId9">${run("open")}</w:hyperlink>${control}</w:p>`
    );
    const whole = select(state, 1, state.doc.child(0).content.size + 1);
    const bare = runCommand(whole, removeLink);

    expect(linkTags(bare, session)).toHaveLength(1);
    expect(bodyXml(bare, session)).toContain(
      `<w:sdtContent><w:hyperlink r:id="rId9">`
    );
  });
});
