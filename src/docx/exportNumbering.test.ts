// @vitest-environment jsdom
import { unzipSync, zipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import {
  AllSelection,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { assert, describe, expect, it } from "vitest";
import {
  bytesEqual,
  decode,
  exportErrorCode,
  fixtureNames,
  makeDeclaredDocx,
  makeDocx,
  makeNumberedDocx,
  ONE_LIST_NUMBERING,
  readFixture,
} from "../__testing__/docx";
import { runCommand } from "../__testing__/editing";
import { canExport } from "../editor/commands/exportQueries";
import {
  toggleBulletList,
  toggleNumberedList,
} from "../editor/commands/listCommands";
import {
  createEditorState,
  editorStateForSession,
} from "../editor/createEditor";
import { richHtmlSlice, withPastedContent } from "../editor/externalClipboard";
import {
  canStartNewList,
  paragraphMarkers,
} from "../editor/plugins/numberingDecorations";
import { toParagraphFormat } from "../model/format";
import { templateList } from "../numbering/listTemplate";
import { parseNumbering } from "../numbering/parseNumbering";
import { R_NS, W_NS } from "../ooxml/xml";
import { exportDocx } from "./exportDocx";
import { importDocx } from "./importDocx";
import { exportProblems } from "./invariants";
import { CONTENT_TYPES_PATH } from "./packageParts";
import type { SessionStore } from "./session";

const NUMBERING_PART = "word/numbering.xml";

const W_NS_DECL = `xmlns:w="${W_NS}"`;

/** The first paragraph that is not a list and has text, plus a position inside it */
function plainParagraph(doc: PMNode): { index: number; pos: number } {
  let spot: { index: number; pos: number } | null = null;
  doc.forEach((block, offset, index) => {
    if (spot !== null || block.type.name !== "paragraph") return;
    if (block.textContent.length === 0) return;
    if (toParagraphFormat(block.attrs.format)?.numbering) return;
    spot = { index, pos: offset + 1 };
  });
  if (spot === null) throw new Error("no paragraph outside a list");
  return spot;
}

function withCaretAt(state: EditorState, at: number): EditorState {
  return state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, at))
  );
}

function open(name: string) {
  const bytes = readFixture(name);
  const { doc, session } = importDocx(bytes);
  return {
    bytes,
    doc,
    session,
    state: editorStateForSession({ doc, session }),
  };
}

function partsOf(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipSync(bytes);
}

/** Exports a document in which a new list has been started */
function exportWithNewList(name: string): {
  bytes: Uint8Array;
  session: SessionStore;
  /** Which block in the body the paragraph that became a list is */
  index: number;
  out: Uint8Array;
} {
  const { bytes, doc, session, state } = open(name);
  const spot = plainParagraph(doc);
  const at = withCaretAt(state, spot.pos);
  let listed = at;
  expect(
    toggleNumberedList(at, (tr) => {
      listed = at.apply(tr);
    })
  ).toBe(true);
  return {
    bytes,
    session,
    index: spot.index,
    out: exportDocx(listed.doc, session),
  };
}

describe("editing that starts a new list", () => {
  it("declares w on an existing numbering part that uses another prefix", () => {
    const original = ONE_LIST_NUMBERING.replaceAll("w:", "n:").replace(
      "xmlns:w=",
      "xmlns:n="
    );
    const opened = importDocx(
      makeNumberedDocx("<w:p><w:r><w:t>Start here</w:t></w:r></w:p>", original)
    );
    const state = editorStateForSession(opened);
    let listed = state;
    expect(
      toggleNumberedList(state, (tr) => {
        listed = state.apply(tr);
      })
    ).toBe(true);
    const bytes = exportDocx(listed.doc, opened.session);
    const xml = decode(partsOf(bytes)[NUMBERING_PART]);
    expect(xml).toContain(
      'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    );
    const imported = importDocx(bytes);
    const ref = toParagraphFormat(
      imported.doc.child(0).attrs.format
    )?.numbering;
    assert(ref);
    expect(
      imported.session.formatting.numbering.lists.get(ref.numId)?.levels.size
    ).toBe(9);
  });

  it.each(fixtureNames)(
    "%s: numbering.xml keeps the original text intact",
    (name) => {
      const { bytes, out } = exportWithNewList(name);
      const before = decode(partsOf(bytes)[NUMBERING_PART]);
      const after = decode(partsOf(out)[NUMBERING_PART]);

      const firstNumAt = before.indexOf("<w:num ");
      const closeAt = before.lastIndexOf("</w:numbering>");
      expect(firstNumAt).toBeGreaterThan(0);

      // The definition is spliced in before the first number and the number at the very end, so the original text survives as three intact chunks
      expect(after.startsWith(before.slice(0, firstNumAt))).toBe(true);
      expect(after).toContain(before.slice(firstNumAt, closeAt));
      expect(after.endsWith(before.slice(closeAt))).toBe(true);
      expect(after.length).toBeGreaterThan(before.length);
    }
  );

  it.each(fixtureNames)(
    "%s: the body regenerates only that paragraph",
    (name) => {
      const { out, session, index } = exportWithNewList(name);
      const documentXml = decode(partsOf(out)[session.mainPartPath]);

      const head =
        session.documentPrefix +
        session.blocks
          .slice(0, index)
          .map((block) => block.xml)
          .join("");
      const tail =
        session.blocks
          .slice(index + 1)
          .map((block) => block.xml)
          .join("") + session.documentSuffix;
      expect(documentXml.startsWith(head)).toBe(true);
      expect(documentXml.endsWith(tail)).toBe(true);
      expect(
        documentXml.slice(head.length, documentXml.length - tail.length)
      ).toContain("<w:numPr>");
    }
  );

  it.each(fixtureNames)(
    "%s: only the body and numbering.xml change",
    (name) => {
      const { bytes, out, session } = exportWithNewList(name);
      const original = partsOf(bytes);
      const exported = partsOf(out);

      for (const key of Object.keys(original)) {
        if (key === session.mainPartPath || key === NUMBERING_PART) continue;
        expect(bytesEqual(exported[key], original[key])).toBe(true);
      }
    }
  );

  it.each(fixtureNames)(
    "%s: reopening it computes the numbers of the new list",
    (name) => {
      const { out } = exportWithNewList(name);
      const again = importDocx(out);
      const numbering = parseNumbering(again.session.numberingXml);
      const markers = paragraphMarkers(again.doc, numbering);

      // The new list is a single paragraph, so the first number shows up as is
      const newList = markers.filter((marker) => marker.text === "1.");
      expect(newList.length).toBeGreaterThan(0);
    }
  );
});

function listParagraph(numId: number, text: string): string {
  return (
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/>` +
    `<w:numId w:val="${numId}"/></w:numPr></w:pPr>` +
    `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  );
}

describe("a list that had no definition from the moment it was opened", () => {
  const bytes = makeNumberedDocx(
    listParagraph(9, "list with no definition"),
    ONE_LIST_NUMBERING
  );

  it("exporting without editing does not disturb numbering.xml", () => {
    const { doc, session } = importDocx(bytes);
    const exported = partsOf(exportDocx(doc, session));
    expect(
      bytesEqual(exported[NUMBERING_PART], partsOf(bytes)[NUMBERING_PART])
    ).toBe(true);
  });

  it("editing the text does not create a new definition", () => {
    const { doc, session } = importDocx(bytes);
    const state = createEditorState(doc);
    const edited = state.apply(state.tr.insertText("edit", 2));
    const exported = partsOf(exportDocx(edited.doc, session));

    expect(
      bytesEqual(exported[NUMBERING_PART], partsOf(bytes)[NUMBERING_PART])
    ).toBe(true);
  });

  it("starting a new list in that document adds only the new definition", () => {
    const withPlain = makeNumberedDocx(
      listParagraph(9, "list with no definition") +
        '<w:p><w:r><w:t xml:space="preserve">plain paragraph</w:t></w:r></w:p>',
      ONE_LIST_NUMBERING
    );
    const { doc, session } = importDocx(withPlain);
    const state = editorStateForSession({ doc, session });
    const at = withCaretAt(state, plainParagraph(doc).pos);
    let listed = at;
    expect(
      toggleNumberedList(at, (tr) => {
        listed = at.apply(tr);
      })
    ).toBe(true);

    const numbering = decode(
      partsOf(exportDocx(listed.doc, session))[NUMBERING_PART]
    );
    const added = parseNumbering(numbering).lists;
    // Only the already-defined 1 and the newly started list are there. 9 still has no definition
    expect(added.has(9)).toBe(false);
    expect(added.size).toBe(2);
  });
});

describe("starting two lists in one document", () => {
  it("the numbered list and the bullet list each get their own definition", () => {
    const bytes = makeNumberedDocx(
      '<w:p><w:r><w:t xml:space="preserve">first</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t xml:space="preserve">second</w:t></w:r></w:p>',
      ONE_LIST_NUMBERING
    );
    const { doc, session } = importDocx(bytes);
    const state = editorStateForSession({ doc, session });

    const first = withCaretAt(state, 1);
    let numbered = first;
    expect(
      toggleNumberedList(first, (tr) => {
        numbered = first.apply(tr);
      })
    ).toBe(true);

    const second = withCaretAt(numbered, numbered.doc.child(0).nodeSize + 1);
    let bulleted = second;
    expect(
      toggleBulletList(second, (tr) => {
        bulleted = second.apply(tr);
      })
    ).toBe(true);

    const out = exportDocx(bulleted.doc, session);
    const lists = parseNumbering(decode(partsOf(out)[NUMBERING_PART])).lists;
    expect(lists.size).toBe(3);

    const again = importDocx(out);
    const markers = paragraphMarkers(
      again.doc,
      parseNumbering(again.session.numberingXml)
    );
    expect(markers.map((marker) => marker.text)).toEqual(["1.", "●"]);
  });
});

const PLAIN_BODY = '<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>';

const NUMBERING_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

describe("a document without numbering.xml", () => {
  /** Starts a bullet list in the first paragraph, through the command the editor offers */
  function bulleted(bytes: Uint8Array) {
    const { doc, session } = importDocx(bytes);
    const at = withCaretAt(editorStateForSession({ doc, session }), 1);
    let listed = at;
    expect(
      toggleBulletList(at, (tr) => {
        listed = at.apply(tr);
      })
    ).toBe(true);
    return { doc: listed.doc, session };
  }

  it("creates numbering.xml, its relationship and its content type for the first list", () => {
    const listed = bulleted(makeDeclaredDocx(PLAIN_BODY));
    expect(exportProblems(listed.doc, listed.session)).toEqual([]);
    expect(canExport(editorStateForSession(listed))).toBe(true);
    const exported = partsOf(exportDocx(listed.doc, listed.session));

    const numbering = decode(exported[NUMBERING_PART]);
    // No byte order mark ahead of the prolog, as no part this export writes from scratch carries one
    expect(exported[NUMBERING_PART][0]).toBe("<".charCodeAt(0));
    expect(
      numbering.startsWith(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="1">`
      )
    ).toBe(true);
    expect(
      numbering.endsWith(
        '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>'
      )
    ).toBe(true);
    expect(parseNumbering(numbering).lists.size).toBe(1);

    expect(decode(exported["word/_rels/document.xml.rels"])).toContain(
      `Type="${R_NS}/numbering" Target="numbering.xml"`
    );
    expect(decode(exported[CONTENT_TYPES_PATH])).toContain(
      `<Override PartName="/word/numbering.xml" ContentType="${NUMBERING_CONTENT_TYPE}"/>`
    );
  });

  it("reopens with the list it defined, drawn from the part it added", () => {
    const listed = bulleted(makeDeclaredDocx(PLAIN_BODY));
    const again = importDocx(exportDocx(listed.doc, listed.session));

    expect(again.session.numberingPartPath).toBe(NUMBERING_PART);
    const markers = paragraphMarkers(
      again.doc,
      parseNumbering(again.session.numberingXml)
    );
    expect(markers.map((marker) => marker.text)).toEqual(["●"]);
  });

  it("fills the missing part an existing numbering relationship names", () => {
    const parts = unzipSync(makeDeclaredDocx(PLAIN_BODY));
    const encoder = new TextEncoder();
    const relsPath = "word/_rels/document.xml.rels";
    parts[relsPath] = encoder.encode(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="${R_NS}/numbering" Target="../lists/definitions.xml"/></Relationships>`
    );
    const listed = bulleted(zipSync(parts));
    const bytes = exportDocx(listed.doc, listed.session);
    const exported = partsOf(bytes);
    expect(exported["lists/definitions.xml"]).toBeDefined();
    expect(bytesEqual(exported[relsPath], parts[relsPath])).toBe(true);
    const reopened = importDocx(bytes);
    expect(reopened.session.numberingPartPath).toBe("lists/definitions.xml");
    expect(
      paragraphMarkers(reopened.doc, reopened.session.formatting.numbering).map(
        (marker) => marker.text
      )
    ).toEqual(["●"]);
  });

  it("takes a name beside a numbering.xml the package holds but relates to nothing", () => {
    const parts = unzipSync(makeDeclaredDocx(PLAIN_BODY));
    const stray = new TextEncoder().encode(`<w:numbering xmlns:w="${W_NS}"/>`);
    parts[NUMBERING_PART] = stray;
    const listed = bulleted(zipSync(parts));
    const exported = partsOf(exportDocx(listed.doc, listed.session));

    expect(bytesEqual(exported[NUMBERING_PART], stray)).toBe(true);
    expect(
      parseNumbering(decode(exported["word/numbering2.xml"])).lists.size
    ).toBe(1);
    expect(decode(exported["word/_rels/document.xml.rels"])).toContain(
      `Type="${R_NS}/numbering" Target="numbering2.xml"`
    );
  });

  it("leaves a package that started no list as it stood", () => {
    const bytes = makeDeclaredDocx(PLAIN_BODY);
    const { doc, session } = importDocx(bytes);
    const exported = partsOf(exportDocx(doc, session));

    expect(exported[NUMBERING_PART]).toBeUndefined();
    for (const [path, original] of Object.entries(partsOf(bytes))) {
      expect(bytesEqual(exported[path], original)).toBe(true);
    }
  });

  it("stops where the package has no content types to declare the part in", () => {
    const opened = importDocx(makeDocx(PLAIN_BODY));
    expect(canStartNewList(editorStateForSession(opened))).toBe(false);

    // The list commands do not apply there, so the list is started the way a plugin would
    const at = withCaretAt(createEditorState(opened.doc), 1);
    let listed = at;
    expect(
      toggleBulletList(at, (tr) => {
        listed = at.apply(tr);
      })
    ).toBe(true);

    expect(exportErrorCode(() => exportDocx(listed.doc, opened.session))).toBe(
      "missing-content-types"
    );
    expect(
      exportProblems(listed.doc, opened.session).map((problem) => problem.code)
    ).toEqual(["missing-content-types"]);
    expect(
      canExport(
        editorStateForSession({ doc: listed.doc, session: opened.session })
      )
    ).toBe(false);
  });
});

/** A numbering part defining a single list under an even number, so the next one is odd */
const EVEN_LIST_NUMBERING =
  `<w:numbering ${W_NS_DECL}>` +
  '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">' +
  '<w:numFmt w:val="bullet"/><w:lvlText w:val="●"/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

/** The state of a document whose only paragraph is plain text, with the caret in it */
function openedAt(bytes: Uint8Array): EditorState {
  const { doc, session } = importDocx(bytes);
  return withCaretAt(editorStateForSession({ doc, session }), 1);
}

describe("the definition a new list is exported with", () => {
  const body = '<w:p><w:r><w:t xml:space="preserve">first</w:t></w:r></w:p>';

  it.each([
    { start: 1e21 },
    { run: { bold: true } },
    {
      indent: {
        startTwips: 720,
        endTwips: null,
        hangingTwips: -1,
        firstLineTwips: null,
      },
    },
  ])("refuses a registered value it cannot preserve (%j)", (invalid) => {
    const opened = importDocx(makeDeclaredDocx(body));
    const listed = runCommand(
      editorStateForSession(opened),
      toggleNumberedList
    );
    const ref = toParagraphFormat(listed.doc.child(0).attrs.format)?.numbering;
    assert(ref);
    const level = templateList("numbered").levels.get(0);
    assert(level);
    const state = listed.apply(
      listed.tr.setDocAttribute("newLists", [
        {
          numId: ref.numId,
          levels: [{ ilvl: 0, ...level, ...invalid }],
        },
      ])
    );
    expect(canExport(state)).toBe(false);
    expect(
      exportProblems(state.doc, opened.session).map((problem) => problem.code)
    ).toEqual(["unsupported-content"]);
    expect(exportErrorCode(() => exportDocx(state.doc, opened.session))).toBe(
      "unsupported-content"
    );
  });

  it("is the one the list was registered with and not one derived from its number", () => {
    const bytes = makeNumberedDocx(body, EVEN_LIST_NUMBERING);
    const { session } = importDocx(bytes);
    const listed = runCommand(openedAt(bytes), toggleNumberedList);

    // The document spends 2, so the list takes 3 - the number an odd/even rule would have read
    // as a bullet list
    const numId = toParagraphFormat(listed.doc.child(0).attrs.format)?.numbering
      ?.numId;
    expect(numId).toBe(3);

    const written = parseNumbering(
      decode(partsOf(exportDocx(listed.doc, session))[NUMBERING_PART])
    );
    expect(written.lists.get(3)?.levels.get(0)?.format).toBe("decimal");
    expect(written.lists.get(3)?.levels.get(0)?.text).toBe("%1.");
  });

  it("two lists of the same kind get different numbers and the very same definition", () => {
    const bytes = makeNumberedDocx(
      body + '<w:p><w:r><w:t xml:space="preserve">second</w:t></w:r></w:p>'
    );
    const { session, doc } = importDocx(bytes);
    const state = editorStateForSession({ doc, session });
    const first = runCommand(withCaretAt(state, 1), toggleBulletList);
    const second = runCommand(
      withCaretAt(first, first.doc.child(0).nodeSize + 1),
      toggleBulletList
    );

    const written = parseNumbering(
      decode(partsOf(exportDocx(second.doc, session))[NUMBERING_PART])
    );
    expect([...written.lists.keys()].sort((a, b) => a - b)).toEqual([1, 2, 3]);
    expect(written.lists.get(2)).toEqual(written.lists.get(3));
  });

  it("a pasted list is exported from its definition the same way a started one is", () => {
    const bytes = makeNumberedDocx(body, EVEN_LIST_NUMBERING);
    const { doc, session } = importDocx(bytes);
    const opened = editorStateForSession({ doc, session });
    const all = opened.apply(
      opened.tr.setSelection(new AllSelection(opened.doc))
    );
    const content = richHtmlSlice(
      all,
      document,
      "<ol><li>One</li><li>Two</li></ol>"
    );
    if (content === null) throw new Error("the markup read as nothing");
    const pasted = all.apply(
      withPastedContent(all.tr.replaceSelection(content.slice), content)
    );

    const written = parseNumbering(
      decode(partsOf(exportDocx(pasted.doc, session))[NUMBERING_PART])
    );
    expect(written.lists.get(3)?.levels.get(0)?.format).toBe("decimal");
    expect(written.lists.get(3)?.levels.get(0)?.text).toBe("%1.");
  });
});

describe("a list number no definition stands behind", () => {
  /**
   * The first paragraph put into a list the way a plugin of a consumer's own would put it there:
   * the reference is written and no definition is registered for it.
   */
  function listedByHand(bytes: Uint8Array): {
    doc: PMNode;
    session: SessionStore;
  } {
    const { doc, session } = importDocx(bytes);
    const state = editorStateForSession({ doc, session });
    const listed = state.apply(
      state.tr.setNodeMarkup(0, undefined, {
        ...doc.child(0).attrs,
        format: { numbering: { numId: 9, ilvl: 0 } },
      })
    );
    return { doc: listed.doc, session };
  }

  const body = '<w:p><w:r><w:t xml:space="preserve">first</w:t></w:r></w:p>';

  it("is refused rather than written into a file whose list is defined nowhere", () => {
    const { doc, session } = listedByHand(makeNumberedDocx(body));

    expect(exportErrorCode(() => exportDocx(doc, session))).toBe(
      "unsupported-content"
    );
  });

  it("is reported ahead of the write, with the message the refusal carries", () => {
    const { doc, session } = listedByHand(makeNumberedDocx(body));

    expect(exportProblems(doc, session)).toEqual([
      {
        code: "unsupported-content",
        message: "the list numbered 9 has no definition to be written",
        pos: 0,
      },
    ]);
  });
});
