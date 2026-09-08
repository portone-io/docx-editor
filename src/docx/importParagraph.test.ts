// @vitest-environment jsdom
import type { Mark, Node as PMNode } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { parseXml, R_NS } from "../ooxml/xml";
import type { LinkTargets } from "./hyperlink";
import { buildParagraph, NO_IMPORT_SOURCES } from "./importParagraph";

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** The single element a fragment holds */
function element(xml: string): Element {
  const wrapped = parseXml(`<w:wrap ${W_NS} xmlns:r="${R_NS}">${xml}</w:wrap>`);
  const el = wrapped.documentElement.firstElementChild;
  if (!el) throw new Error("no element");
  return el;
}

/** Moves a single XML fragment into a paragraph node */
function requireParagraph(xml: string, links?: LinkTargets): PMNode {
  return buildParagraph(
    element(xml),
    null,
    links ? { ...NO_IMPORT_SOURCES, links } : NO_IMPORT_SOURCES
  );
}

/** The one preserved fragment a paragraph holding nothing else came out as */
function lonePreserved(xml: string, links?: LinkTargets): PMNode {
  const node = requireParagraph(xml, links);
  expect(node.type.name).toBe("paragraph");
  expect(node.childCount).toBe(1);
  return node.child(0);
}

const run = (text: string, rPr = "") => `<w:r>${rPr}<w:t>${text}</w:t></w:r>`;

const ID_PR = '<w:sdtPr><w:id w:val="7"/></w:sdtPr>';

const sdt = (inner: string, pr = ID_PR) =>
  `<w:sdt>${pr}<w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** The marks one inline node wears, in the order they are drawn */
function markNames(node: PMNode): string[] {
  return node.marks.map((mark) => mark.type.name);
}

function requireSdtMark(node: PMNode): Mark {
  const mark = node.marks.find((entry) => entry.type.name === "sdt");
  if (!mark) throw new Error("no sdt mark");
  return mark;
}

function requireLinkMark(node: PMNode): Mark {
  const mark = node.marks.find((entry) => entry.type.name === "link");
  if (!mark) throw new Error("no link mark");
  return mark;
}

describe("a stretch of text wrapped in a content control", () => {
  const LOCK_PR =
    '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

  it("stays editable text, wearing the control's mark outside its run", () => {
    const node = requireParagraph(
      `<w:p>${sdt(run("2026-08-04"), LOCK_PR)}</w:p>`
    );
    expect(node.childCount).toBe(1);

    const child = node.child(0);
    expect(child.isText).toBe(true);
    expect(child.text).toBe("2026-08-04");
    // The control is declared ahead of the run, so it is the outer of the two spans on screen
    expect(markNames(child)).toEqual(["sdt", "run"]);
    expect(requireSdtMark(child).attrs).toEqual({
      sdtPrefix: `<w:sdt>${LOCK_PR}`,
      sdtKey: 0,
      contentsLocked: true,
      deletionLocked: true,
    });
  });

  it("a control that does not lock its contents comes in unlocked", () => {
    const node = requireParagraph(`<w:p>${sdt(run("value"))}</w:p>`);
    expect(requireSdtMark(node.child(0)).attrs).toEqual({
      sdtPrefix: `<w:sdt>${ID_PR}`,
      sdtKey: 0,
      contentsLocked: false,
      deletionLocked: false,
    });
  });

  /** The two clauses each of the four values settles, which are independent of one another */
  it.each([
    ["contentLocked", true, false],
    ["sdtContentLocked", true, true],
    ["sdtLocked", false, true],
    ["unlocked", false, false],
  ])(
    "a lock of %s reads as contents locked: %s, deletion locked: %s",
    (val, contentsLocked, deletionLocked) => {
      const node = requireParagraph(
        `<w:p>${sdt(run("value"), `<w:sdtPr><w:lock w:val="${val}"/></w:sdtPr>`)}</w:p>`
      );
      const attrs = requireSdtMark(node.child(0)).attrs;
      expect(attrs.contentsLocked).toBe(contentsLocked);
      expect(attrs.deletionLocked).toBe(deletionLocked);
    }
  );

  it("the wrapper's own attributes and its sdtEndPr come along", () => {
    const prefix =
      `<w:sdt id="7">${ID_PR}` + "<w:sdtEndPr><w:rPr/></w:sdtEndPr>";
    const node = requireParagraph(
      `<w:p>${prefix}<w:sdtContent>${run("value")}</w:sdtContent></w:sdt></w:p>`
    );
    expect(requireSdtMark(node.child(0)).attrs.sdtPrefix).toBe(prefix);
  });

  it("several runs inside one control keep their formatting and share the one mark", () => {
    const node = requireParagraph(
      `<w:p>${sdt(run("2026") + run("AD", "<w:rPr><w:b/></w:rPr>"))}</w:p>`
    );
    expect(node.childCount).toBe(2);
    expect(node.textContent).toBe("2026AD");
    expect(
      requireSdtMark(node.child(0)).eq(requireSdtMark(node.child(1)))
    ).toBe(true);
    expect(node.child(1).marks.at(-1)?.attrs.rPr).toBe("<w:rPr><w:b/></w:rPr>");
  });

  it("the runs on either side of the control wear no mark of it", () => {
    const node = requireParagraph(
      `<w:p>${run("signed on ")}${sdt(run("2026-08-04"))}${run(" final")}</w:p>`
    );
    expect(node.children.map(markNames)).toEqual([
      ["run"],
      ["sdt", "run"],
      ["run"],
    ]);
  });

  it("a run with no text inside the control is preserved wearing the mark too", () => {
    const empty = '<w:r><w:rPr><w:rtl w:val="0"/></w:rPr></w:r>';
    const node = requireParagraph(`<w:p>${sdt(empty + run("value"))}</w:p>`);

    expect(node.child(0).type.name).toBe("rawInline");
    expect(node.child(0).attrs.xml).toBe(empty);
    // The control has to close around this XML again on export
    expect(markNames(node.child(0))).toEqual(["sdt"]);
  });

  it("a break and a tab inside the control wear the mark as well", () => {
    const node = requireParagraph(
      `<w:p>${sdt("<w:r><w:tab/><w:br/></w:r>")}</w:p>`
    );
    expect(node.children.map((child) => child.type.name)).toEqual([
      "text",
      "hardBreak",
    ]);
    expect(node.child(0).text).toBe("\t");
    expect(node.children.map(markNames)).toEqual([
      ["sdt", "run", "tab"],
      ["sdt", "run"],
    ]);
  });

  it("keeps the source attributes on a tab mark", () => {
    const node = requireParagraph(
      '<w:p><w:r><w:tab w:custom="keep"/></w:r></w:p>'
    );
    expect(node.textContent).toBe("\t");
    expect(
      node.child(0).marks.find((mark) => mark.type.name === "tab")?.attrs
    ).toEqual({ tabAttrs: 'w:custom="keep"' });
  });
});

describe("two neighbouring controls written exactly alike", () => {
  const twins = `<w:p>${sdt(run("a"))}${sdt(run("b"))}</w:p>`;

  it("stay two stretches of their own rather than running into one", () => {
    const node = requireParagraph(twins);
    expect(node.childCount).toBe(2);
    expect(node.children.map((child) => child.text)).toEqual(["a", "b"]);
  });

  it("wear a mark each, telling them apart by nothing but the control they are", () => {
    const node = requireParagraph(twins);
    const first = requireSdtMark(node.child(0));
    const second = requireSdtMark(node.child(1));

    expect(first.eq(second)).toBe(false);
    expect(first.attrs.sdtPrefix).toBe(second.attrs.sdtPrefix);
    expect([first.attrs.sdtKey, second.attrs.sdtKey]).toEqual([0, 1]);
  });
});

describe("a control holding something that puts nothing on screen", () => {
  const PROOF_ERR = '<w:proofErr w:type="spellStart"/>';
  const BOOKMARK_START = '<w:bookmarkStart w:id="1" w:name="signedOn"/>';
  const BOOKMARK_END = '<w:bookmarkEnd w:id="1"/>';
  const LOCK_PR =
    '<w:sdtPr><w:id w:val="7"/><w:lock w:val="sdtContentLocked"/></w:sdtPr>';

  const body = `<w:p>${sdt(
    PROOF_ERR + BOOKMARK_START + run("2026-08-04") + BOOKMARK_END,
    LOCK_PR
  )}</w:p>`;

  it("stays a paragraph to edit, with that XML kept right where it was", () => {
    const node = requireParagraph(body);
    expect(node.children.map((child) => child.type.name)).toEqual([
      "rawInline",
      "rawInline",
      "text",
      "rawInline",
    ]);
    expect(node.children.map((child) => child.attrs.xml)).toEqual([
      PROOF_ERR,
      BOOKMARK_START,
      undefined,
      BOOKMARK_END,
    ]);
    expect(node.textContent).toBe("2026-08-04");
  });

  it("holds the whole stretch inside the one control, lock and all", () => {
    const node = requireParagraph(body);
    expect(node.children.map(markNames)).toEqual([
      ["sdt"],
      ["sdt"],
      ["sdt", "run"],
      ["sdt"],
    ]);
    const marks = node.children.map(requireSdtMark);
    expect(marks.every((mark) => mark.eq(marks[0]))).toBe(true);
    expect(marks[0].attrs.contentsLocked).toBe(true);
    expect(marks[0].attrs.deletionLocked).toBe(true);
  });

  it("is still a control with nothing but that XML inside it", () => {
    const node = requireParagraph(`<w:p>${sdt(BOOKMARK_START)}</w:p>`);
    expect(node.childCount).toBe(1);
    expect(markNames(node.child(0))).toEqual(["sdt"]);
  });
});

describe("a control we could not take apart stays one preserved inline", () => {
  it.each([
    [
      "an sdt with no sdtPr",
      `<w:p><w:sdt><w:sdtContent>${run("value")}</w:sdtContent></w:sdt></w:p>`,
    ],
    [
      "an sdtContent carrying an attribute of its own",
      `<w:p><w:sdt>${ID_PR}<w:sdtContent w:val="1">${run("value")}` +
        "</w:sdtContent></w:sdt></w:p>",
    ],
    [
      "an sdt with another element behind its sdtContent",
      `<w:p><w:sdt>${ID_PR}<w:sdtContent>${run("value")}</w:sdtContent>` +
        "<w:sdtEndPr/></w:sdt></w:p>",
    ],
    [
      "a control holding another control",
      `<w:p>${sdt(sdt(run("value")))}</w:p>`,
    ],
    ["a control with nothing inside it at all", `<w:p>${sdt("")}</w:p>`],
  ])("%s", (_name, xml) => {
    const preserved = lonePreserved(xml);

    expect(preserved.type.name).toBe("rawInline");
    expect(preserved.attrs.element).toBe("sdt");
    expect(preserved.attrs.display).toBe("chip");
    expect(preserved.attrs.guarded).toBe(false);
    expect(String(preserved.attrs.xml).startsWith("<w:sdt")).toBe(true);
  });

  it("shows the text that control held, so the chip says what it stands for", () => {
    const preserved = lonePreserved(
      `<w:p><w:sdt><w:sdtContent>${run("value")}</w:sdtContent></w:sdt></w:p>`
    );

    expect(preserved.attrs.text).toBe("value");
  });
});

it("keeps a footnote reference inside a content control", () => {
  const node = requireParagraph(
    `<w:p>${sdt('<w:r><w:footnoteReference w:id="1"/></w:r>')}</w:p>`
  );

  expect(node.child(0).type.name).toBe("noteReference");
  expect(markNames(node.child(0))).toEqual(["sdt", "run"]);
});

describe("a stretch of text wrapped in a hyperlink", () => {
  const EXTERNAL =
    '<w:hyperlink r:id="rId9" w:history="1" w:tooltip="Our terms">';
  const LINKS: LinkTargets = new Map([["rId9", "https://example.com/terms"]]);

  it("stays editable text, wearing the link's mark outside its run", () => {
    const node = requireParagraph(
      `<w:p>${EXTERNAL}${run("our terms")}</w:hyperlink></w:p>`,
      LINKS
    );
    expect(node.childCount).toBe(1);

    const child = node.child(0);
    expect(child.text).toBe("our terms");
    // The link is declared ahead of the run, so it is the outer of the two spans on screen
    expect(markNames(child)).toEqual(["link", "run"]);
  });

  it("carries the address the relationship names, and the opening tag as it came", () => {
    const node = requireParagraph(
      `<w:p>${EXTERNAL}${run("our terms")}</w:hyperlink></w:p>`,
      LINKS
    );
    const mark = requireLinkMark(node.child(0));
    expect(mark.attrs.href).toBe("https://example.com/terms");
    expect(mark.attrs.linkPrefix).toBe(EXTERNAL);
  });

  it("holds the several runs it wrapped inside the one link", () => {
    const node = requireParagraph(
      `<w:p>${EXTERNAL}${run("our ")}${run("terms", "<w:rPr><w:b/></w:rPr>")}` +
        "</w:hyperlink></w:p>",
      LINKS
    );
    const marks = node.children.map(requireLinkMark);
    expect(node.textContent).toBe("our terms");
    expect(marks.every((mark) => mark.eq(marks[0]))).toBe(true);
  });

  it("offers no address where the relationship leads nowhere we follow", () => {
    const node = requireParagraph(
      `<w:p>${EXTERNAL}${run("our terms")}</w:hyperlink></w:p>`
    );
    expect(requireLinkMark(node.child(0)).attrs.href).toBeNull();
    expect(requireLinkMark(node.child(0)).attrs.linkPrefix).toBe(EXTERNAL);
  });

  it("keeps a link naming a bookmark alone, with no address to offer", () => {
    const opening = '<w:hyperlink w:anchor="chapter3">';
    const node = requireParagraph(
      `<w:p>${opening}${run("Chapter Three")}</w:hyperlink></w:p>`,
      LINKS
    );
    const mark = requireLinkMark(node.child(0));
    expect(node.child(0).text).toBe("Chapter Three");
    expect(mark.attrs.href).toBeNull();
    expect(mark.attrs.linkPrefix).toBe(opening);
  });

  it("keeps two links written exactly alike apart", () => {
    const link = `${EXTERNAL}${run("terms")}</w:hyperlink>`;
    const node = requireParagraph(`<w:p>${link}${link}</w:p>`, LINKS);
    const first = requireLinkMark(node.child(0));
    const second = requireLinkMark(node.child(1));

    expect(node.childCount).toBe(2);
    expect(first.eq(second)).toBe(false);
    expect([first.attrs.linkKey, second.attrs.linkKey]).toEqual([0, 1]);
  });

  it("keeps the markup that puts nothing on screen inside the link", () => {
    const bookmark = '<w:bookmarkStart w:id="1" w:name="terms"/>';
    const node = requireParagraph(
      `<w:p>${EXTERNAL}${bookmark}${run("terms")}</w:hyperlink></w:p>`,
      LINKS
    );
    expect(node.children.map((child) => child.type.name)).toEqual([
      "rawInline",
      "text",
    ]);
    expect(node.children.map(markNames)).toEqual([["link"], ["link", "run"]]);
  });
});

describe("a link and a control standing one inside the other", () => {
  const LINKS: LinkTargets = new Map([["rId9", "https://example.com"]]);
  const link = (inner: string) =>
    `<w:hyperlink r:id="rId9">${inner}</w:hyperlink>`;

  it("a control holding a link opens with both marks, the control outside", () => {
    const node = requireParagraph(
      `<w:p>${sdt(link(run("terms")))}</w:p>`,
      LINKS
    );
    expect(markNames(node.child(0))).toEqual(["sdt", "link", "run"]);
    expect(requireLinkMark(node.child(0)).attrs.href).toBe(
      "https://example.com"
    );
  });

  /** The marks record the control outside the link, so the other nesting has nowhere to go */
  it("a link holding a control stays one preserved inline", () => {
    const preserved = lonePreserved(
      `<w:p>${link(sdt(run("terms")))}</w:p>`,
      LINKS
    );

    expect(preserved.type.name).toBe("rawInline");
    expect(preserved.attrs.element).toBe("sdt");
    expect(preserved.attrs.text).toBe("terms");
    // The link around it was read, so the control that stayed whole still wears its mark
    expect(markNames(preserved)).toEqual(["link"]);
  });
});

describe("a link we could not take apart stays one preserved inline", () => {
  it.each([
    [
      "a link with nothing inside it at all",
      '<w:p><w:hyperlink r:id="rId9"></w:hyperlink></w:p>',
    ],
    [
      "a link naming its relationship under another prefix",
      `<w:p><w:hyperlink xmlns:rel="${R_NS}" rel:id="rId9">${run("terms")}` +
        "</w:hyperlink></w:p>",
    ],
  ])("%s", (_name, xml) => {
    const preserved = lonePreserved(xml);

    expect(preserved.type.name).toBe("rawInline");
    expect(preserved.attrs.element).toBe("hyperlink");
    expect(preserved.attrs.display).toBe("chip");
  });
});

it("keeps a footnote reference inside a hyperlink", () => {
  const links: LinkTargets = new Map([["rId9", "https://example.com"]]);
  const node = requireParagraph(
    '<w:p><w:hyperlink r:id="rId9"><w:r><w:footnoteReference w:id="1"/>' +
      "</w:r></w:hyperlink></w:p>",
    links
  );

  expect(node.child(0).type.name).toBe("noteReference");
  expect(markNames(node.child(0))).toEqual(["link", "run"]);
});

describe("run children the editor does not model", () => {
  /** The children of the paragraph a fragment inside one run came out as */
  function runChildren(inner: string): PMNode[] {
    const node = requireParagraph(`<w:p><w:r>${inner}</w:r></w:p>`);
    const children: PMNode[] = [];
    node.forEach((child) => {
      children.push(child);
    });
    return children;
  }

  it("a lastRenderedPageBreak stays inside its run and puts nothing on screen", () => {
    const node = requireParagraph(
      "<w:p><w:r><w:lastRenderedPageBreak/><w:t>a</w:t></w:r></w:p>"
    );

    expect(node.type.name).toBe("paragraph");
    expect(node.textContent).toBe("a");
    expect(node.child(0).type.name).toBe("rawRunContent");
    expect(node.child(0).attrs.display).toBe("hidden");
    expect(node.child(0).attrs.guarded).toBe(false);
    expect(node.child(0).attrs.xml).toBe("<w:lastRenderedPageBreak/>");
  });

  it("a soft hyphen and a no-break hyphen keep their run formatting and draw their characters", () => {
    const children = runChildren(
      "<w:rPr><w:b/></w:rPr><w:softHyphen/><w:noBreakHyphen/>"
    );

    expect(children.map((child) => child.attrs.display)).toEqual([
      "hidden",
      "text",
    ]);
    expect(children[1].attrs.text).toBe("‑");
    for (const child of children) {
      expect(markNames(child)).toEqual(["run"]);
      expect(child.marks[0].attrs.rPr).toBe("<w:rPr><w:b/></w:rPr>");
    }
  });

  it("a carriage return draws a line break and goes back out as w:cr", () => {
    const [child] = runChildren("<w:cr/>");

    expect(child.type.name).toBe("rawRunContent");
    expect(child.attrs.display).toBe("break");
    expect(child.attrs.xml).toBe("<w:cr/>");
  });

  it("a field character run stays a chip that wears the run mark", () => {
    const children = runChildren(
      '<w:fldChar w:fldCharType="begin"/>' +
        '<w:instrText xml:space="preserve"> PAGE </w:instrText>'
    );

    expect(children.map((child) => child.type.name)).toEqual([
      "rawRunContent",
      "rawRunContent",
    ]);
    expect(children.map((child) => child.attrs.element)).toEqual([
      "fldChar",
      "instrText",
    ]);
    // A field falls apart if a piece of it goes, so the guard answers for every one of them
    expect(children.map((child) => child.attrs.guarded)).toEqual([true, true]);
    expect(children[1].attrs.text).toBe(" PAGE ");
    expect(markNames(children[0])).toEqual(["run"]);
  });

  it("a drawing we cannot read stays a chip instead of preserving the paragraph", () => {
    const [child] = runChildren(
      '<w:drawing><wp:anchor xmlns:wp="urn:x"/></w:drawing>'
    );

    expect(child.type.name).toBe("rawRunContent");
    expect(child.attrs.element).toBe("drawing");
    expect(child.attrs.display).toBe("chip");
  });
});

describe("paragraph children the editor does not model", () => {
  it("a revision container is a visible chip carrying its text", () => {
    const node = requireParagraph(
      `<w:p>${run("a")}<w:ins w:id="1" w:author="Reviewer A" ` +
        `w:date="2026-01-01T00:00:00Z">${run("inserted")}</w:ins>${run("b")}</w:p>`
    );

    expect(node.textContent).toBe("ab");
    expect(node.child(1).type.name).toBe("rawInline");
    expect(node.child(1).attrs.element).toBe("ins");
    expect(node.child(1).attrs.display).toBe("chip");
    expect(node.child(1).attrs.text).toBe("inserted");
    expect(node.child(1).attrs.guarded).toBe(false);
  });

  it("a simple field is a visible chip carrying its cached result", () => {
    const preserved = lonePreserved(
      `<w:p><w:fldSimple w:instr=" PAGE ">${run("7")}</w:fldSimple></w:p>`
    );

    expect(preserved.attrs.element).toBe("fldSimple");
    expect(preserved.attrs.display).toBe("chip");
    expect(preserved.attrs.text).toBe("7");
  });

  it("a range marker is invisible and the guard answers for it", () => {
    const preserved = lonePreserved(
      '<w:p><w:permStart w:id="3" w:edGrp="everyone"/></w:p>'
    );

    expect(preserved.attrs.element).toBe("permStart");
    expect(preserved.attrs.display).toBe("hidden");
    expect(preserved.attrs.guarded).toBe(true);
  });

  it("a proofing mark is invisible and nothing depends on it", () => {
    const preserved = lonePreserved(
      '<w:p><w:proofErr w:type="spellStart"/></w:p>'
    );

    expect(preserved.attrs.display).toBe("hidden");
    expect(preserved.attrs.guarded).toBe(false);
  });

  it("a run with nothing in it is still kept whole and says so", () => {
    const preserved = lonePreserved(
      "<w:p><w:r><w:rPr><w:b/></w:rPr></w:r></w:p>"
    );

    expect(preserved.type.name).toBe("rawInline");
    expect(preserved.attrs.element).toBe("r");
    expect(preserved.attrs.display).toBe("hidden");
  });
});
