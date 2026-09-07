/**
 * The battery of edits that runs before an export is validated, as one probe per public writer.
 *
 * Writing a paragraph back reaches the paragraph serializer and nothing else: everything the
 * export writes only for content that was not in the file already - a list definition spliced into
 * numbering.xml, a table built from the template, a media part with its relationship and content
 * type, a hyperlink with the external relationship its address lives on, a content control around
 * a locked stretch, a comment with the parts that hang off it - stays out of reach until a command
 * puts it there. A probe is one command of the public surface (`./commands`, `./table`) run over
 * an `EditorState` with no view, and the battery is every one of them in turn.
 *
 * `WRITER_PROBES` is keyed by the export the probe runs, and `NOT_A_WRITER` carries every other
 * export with the reason it reaches no writer. `docx/writerProbes.test.ts` fails the moment an
 * export appears in neither, which is what stops a writer from being added and never validated.
 *
 * The order the probes are written in is the order they run in: a probe works on what the ones
 * before it left, and the table probes need the table an earlier one inserted.
 */

import { unzipSync } from "fflate";
import type { Node as PMNode } from "prosemirror-model";
import {
  type Command,
  type EditorState,
  TextSelection,
} from "prosemirror-state";
import { CellSelection, TableMap } from "prosemirror-tables";
import { expect } from "vitest";
import {
  bytesEqual,
  decodeBase64,
  TINY_PNG,
  TINY_PNG_DATA_URL,
} from "../../__testing__/docx";
import {
  type ActiveParagraphAlign,
  activeLineSpacing,
  activeParagraphAlign,
  addComment,
  type ImageToInsert,
  insertImage,
  insertTable,
  lockSelection,
  setLineSpacing,
  setLink,
  setParagraphAlign,
  setTextColor,
  toggleBold,
  toggleBulletList,
  toggleNumberedList,
  unlockSelection,
} from "../../editor/commands";
import { createEditorState } from "../../editor/createEditor";
import {
  type LineSpacing,
  type ParagraphAlign,
  toParagraphFormat,
} from "../../model/format";
import { wAttr } from "../../ooxml/units";
import { childByLocalName, decodeUtf8, parseXml, W_NS } from "../../ooxml/xml";
import {
  addRowAfter,
  mergeCells,
  setCellPadding,
  setCellVerticalAlign,
} from "../../table";
import { exportDocx } from "../exportDocx";
import { documentNumbering, type SessionStore } from "../session";

/** The editing state a screen would hold for this document */
export function openState(doc: PMNode, session: SessionStore): EditorState {
  return createEditorState(doc, {
    numbering: documentNumbering(session),
    styles: session.styles,
    defaults: session.defaults,
    canStartNewList: session.numberingPartPath !== null,
    paragraphStyles: session.paragraphStyles,
  });
}

/**
 * Runs one command over the selection the state holds, and refuses to go on when it reported that
 * it does not apply or when dispatching it changed nothing. Which probe was running is what the
 * battery puts in front of the message.
 */
export function ran(state: EditorState, command: Command): EditorState {
  let next = state;
  const handled = command(state, (tr) => {
    next = state.apply(tr);
  });
  if (!handled) throw new Error("the command reports that it does not apply");
  if (next === state) throw new Error("the command dispatched nothing");
  return next;
}

interface Spot {
  pos: number;
  node: PMNode;
}

/** The body paragraphs holding text, which is the list a slot names a place in */
function textParagraphs(doc: PMNode): Spot[] {
  const spots: Spot[] = [];
  doc.forEach((block, pos) => {
    if (block.type.name === "paragraph" && block.textContent !== "") {
      spots.push({ pos, node: block });
    }
  });
  return spots;
}

/**
 * The paragraphs the battery works in, each named for what is done to it there.
 *
 * A probe says which one it stands in and the battery puts the caret there before running it, so
 * that no probe depends on where the one before it left the selection. `table` is the table the
 * `tabled` probe inserted, which is where the table probes work.
 */
export type ProbeSlot =
  | "formatted"
  | "spaced"
  | "numbered"
  | "bulleted"
  | "tabled"
  | "pictured"
  | "table";

const PARAGRAPH_SLOTS = [
  "formatted",
  "spaced",
  "numbered",
  "bulleted",
  "tabled",
  "pictured",
] as const;

type ParagraphSlot = (typeof PARAGRAPH_SLOTS)[number];

type SlotPlaces = Readonly<Record<ParagraphSlot, number>>;

/**
 * Which paragraph each slot names, as its place among the body paragraphs holding text.
 *
 * The places are read off the document the battery starts from, and a slot is resolved against
 * the current document every time it is used. That list is what none of the probes moves: they
 * neither empty a paragraph nor add one holding text, and the empty paragraph an inserted table
 * leaves behind is not in it. A list marker is: a paragraph the battery numbers stops being plain,
 * so the reservation below is made once, where every paragraph is still as the document had it.
 */
function slotPlaces(doc: PMNode): SlotPlaces {
  const plain = textParagraphs(doc).flatMap(({ node }, index) =>
    toParagraphFormat(node.attrs.format)?.numbering === undefined ? [index] : []
  );
  const places: Partial<Record<ParagraphSlot, number>> = {};
  PARAGRAPH_SLOTS.forEach((slot, at) => {
    const place = plain[at];
    if (place === undefined) {
      throw new Error(
        `the battery reserves ${PARAGRAPH_SLOTS.length} plain paragraphs and the document has ${plain.length}`
      );
    }
    places[slot] = place;
  });
  return places as SlotPlaces;
}

function paragraphAt(doc: PMNode, place: number): Spot {
  const spot = textParagraphs(doc)[place];
  if (spot === undefined) {
    throw new Error(`the document has no paragraph holding text at ${place}`);
  }
  return spot;
}

/** The table the battery inserted, which stands right after the paragraph the caret was in */
function tableAfter(doc: PMNode, paragraph: Spot): Spot {
  const $after = doc.resolve(paragraph.pos + paragraph.node.nodeSize);
  const table = $after.nodeAfter;
  if (table === null || table.type.spec.tableRole !== "table") {
    throw new Error("no table stands where the battery inserted one");
  }
  return { pos: $after.pos, node: table };
}

function caretAt(state: EditorState, pos: number): EditorState {
  return state.apply(
    state.tr.setSelection(TextSelection.near(state.doc.resolve(pos)))
  );
}

/** The caret where the slot names, which is what a probe is handed */
function placedIn(
  state: EditorState,
  slot: ProbeSlot,
  places: SlotPlaces
): EditorState {
  const paragraph = paragraphAt(
    state.doc,
    places[slot === "table" ? "tabled" : slot]
  );
  return slot === "table"
    ? caretAt(state, tableAfter(state.doc, paragraph).pos + 1)
    : caretAt(state, paragraph.pos + 1);
}

/** The whole of the paragraph the caret stands in, which is what a formatting command takes */
export function wholeParagraph(state: EditorState): EditorState {
  const $from = state.selection.$from;
  return state.apply(
    state.tr.setSelection(
      TextSelection.create(state.doc, $from.start(), $from.end())
    )
  );
}

/** The table the selection sits in */
function tableAround(state: EditorState): Spot {
  const $from = state.selection.$from;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    const node = $from.node(depth);
    if (node.type.spec.tableRole === "table") {
      return { pos: $from.before(depth), node };
    }
  }
  throw new Error("the selection is not in a table");
}

/** Two neighbouring cells of one row of that table selected, which is what a merge takes */
export function twoCellsOfRow(state: EditorState, row: number): EditorState {
  const table = tableAround(state);
  const map = TableMap.get(table.node);
  const start = table.pos + 1;
  return state.apply(
    state.tr.setSelection(
      CellSelection.create(
        state.doc,
        start + map.positionAt(row, 0, table.node),
        start + map.positionAt(row, 1, table.node)
      )
    )
  );
}

const ALIGNS: readonly ParagraphAlign[] = [
  "center",
  "right",
  "justify",
  "left",
];

/**
 * An alignment the paragraph is not already drawn with, and a line spacing it does not already
 * carry. Both commands leave a paragraph that already reads that way untouched, so a probe asking
 * for the value it already has would report it changed nothing.
 */
function otherAlign(active: ActiveParagraphAlign): ParagraphAlign {
  return (
    ALIGNS.find((align) => active.kind === "mixed" || align !== active.align) ??
    "center"
  );
}

function otherSpacing(active: LineSpacing | null): LineSpacing {
  const doubled = active?.rule === "auto" && active.lines === 2;
  return { rule: "auto", lines: doubled ? 1.5 : 2 };
}

const TEXT_COLOR = "#1F4E79";

/** The address a probe links a stretch of text to, which the export writes a relationship for */
const LINK_ADDRESS = "https://example.com/battery?a=1&b=2";

const CONTENT_TYPES_PATH = "[Content_Types].xml";

const A_PICTURE: ImageToInsert = {
  src: TINY_PNG_DATA_URL,
  extent: { cx: 952500, cy: 952500 },
  alt: "the picture a probe inserted",
};

/** A 1x1 transparent GIF, an image of a kind no fixture's content types declare */
const TINY_GIF_BASE64 =
  "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";

const TINY_GIF = decodeBase64(TINY_GIF_BASE64);

/**
 * A second picture of that other kind, so the export has to add the declaration rather than find
 * it there already, which is the only way the content types writer runs at all
 */
const ANOTHER_PICTURE: ImageToInsert = {
  src: `data:image/gif;base64,${TINY_GIF_BASE64}`,
  extent: { cx: 476250, cy: 476250 },
  alt: "the second picture a probe inserted",
};

/** A package the export wrote, and the document it was written from */
export interface ExportedPackage {
  /** The fixture or document the package was written for, which the failures are named by */
  readonly name: string;
  readonly session: SessionStore;
  readonly bytes: Uint8Array;
  readonly parts: Readonly<Record<string, Uint8Array>>;
  /** The text of one part, which the package has to hold */
  text(path: string): string;
  /** The text of the main story */
  readonly mainXml: string;
}

export function exportedPackage(
  name: string,
  doc: PMNode,
  session: SessionStore
): ExportedPackage {
  const bytes = exportDocx(doc, session);
  const parts = unzipSync(bytes);
  const text = (path: string): string => {
    const part = parts[path];
    if (part === undefined) throw new Error(`${name} wrote no ${path}`);
    return decodeUtf8(part).text;
  };
  return {
    name,
    session,
    bytes,
    parts,
    text,
    mainXml: text(session.mainPartPath),
  };
}

/** The part as the document was opened with it, which is what an addition is measured against */
function openedPart(exported: ExportedPackage, path: string): Uint8Array {
  const part = exported.session.parts.get(path);
  if (part === undefined) {
    throw new Error(`${exported.name} carries no ${path}`);
  }
  return part;
}

function mediaPaths(exported: ExportedPackage): string[] {
  return Object.keys(exported.parts).filter((path) => path.includes("/media/"));
}

/** The media parts the probes added, which the same document written out without them has not */
function addedMedia(
  exported: ExportedPackage,
  untouched: ExportedPackage
): string[] {
  const before = mediaPaths(untouched);
  return mediaPaths(exported).filter((path) => !before.includes(path));
}

/** Whether a relationships part of the package points at that media file */
function pointsAt(exported: ExportedPackage, mediaPath: string): boolean {
  const target = `media/${mediaPath.slice(mediaPath.lastIndexOf("/") + 1)}`;
  return Object.entries(exported.parts).some(
    ([path, bytes]) =>
      path.endsWith(".rels") && decodeUtf8(bytes).text.includes(target)
  );
}

/** The one media part a probe wrote for an inserted image, carrying its bytes and a relationship */
function expectMediaPart(
  exported: ExportedPackage,
  added: readonly string[],
  extension: string,
  bytes: Uint8Array
): void {
  const written = added.filter((path) => path.endsWith(`.${extension}`));
  expect(written, `${exported.name} .${extension} media parts`).toHaveLength(1);

  const path = written[0];
  expect(
    bytesEqual(exported.parts[path], bytes),
    `${exported.name} ${path}`
  ).toBe(true);
  expect(
    pointsAt(exported, path),
    `${exported.name} ${path} relationship`
  ).toBe(true);
}

/** Every hyperlink relationship the package holds, wherever it holds it */
function linkRelationships(exported: ExportedPackage): string[] {
  return Object.entries(exported.parts).flatMap(([path, bytes]) =>
    path.endsWith(".rels")
      ? (decodeUtf8(bytes).text.match(/<Relationship[^>]*hyperlink[^>]*\/>/g) ??
        [])
      : []
  );
}

function numbersOf(values: readonly (string | null)[]): number[] {
  return values.flatMap((value) => {
    const parsed = Number.parseInt(value ?? "", 10);
    return Number.isFinite(parsed) ? [parsed] : [];
  });
}

/** The list numbers numbering.xml defines */
function definedNumIds(xml: string): Set<number> {
  const nums = Array.from(parseXml(xml).getElementsByTagNameNS(W_NS, "num"));
  return new Set(numbersOf(nums.map((num) => wAttr(num, "numId"))));
}

/** The list numbers the body's paragraphs point at */
function referencedNumIds(xml: string): Set<number> {
  const refs = Array.from(parseXml(xml).getElementsByTagNameNS(W_NS, "numPr"));
  return new Set(
    numbersOf(
      refs.map((ref) => {
        const numId = childByLocalName(ref, "numId");
        return numId === null ? null : wAttr(numId, "val");
      })
    )
  );
}

/** The list definitions the probes added to the numbering the document opened with */
function addedNumIds(exported: ExportedPackage): number[] {
  const { numberingPartPath, numberingXml } = exported.session;
  if (numberingPartPath === null || numberingXml === null) {
    throw new Error(
      `${exported.name} carries no numbering.xml to define a list in`
    );
  }
  const defined = definedNumIds(numberingXml);
  return Array.from(definedNumIds(exported.text(numberingPartPath))).filter(
    (numId) => !defined.has(numId)
  );
}

function lockedControlCount(xml: string): number {
  return xml.match(/sdtContentLocked/g)?.length ?? 0;
}

/**
 * One command of the public surface, run where it writes something the export has to build.
 *
 * `expect` is what the probe says of the package afterwards, measured against the same document
 * written out without the battery because a fixture may carry a picture or a link of its own. A
 * probe that adds nothing to the package beyond markup the schemas already judge leaves it out:
 * running at all is what it asserts, since the battery refuses to go on when a command reports it
 * changed nothing.
 */
export interface WriterProbe {
  /** What the probe does, which is what a battery that could not run it is reported by */
  name: string;
  slot: ProbeSlot;
  run(state: EditorState): EditorState;
  expect?(exported: ExportedPackage, untouched: ExportedPackage): void;
}

export const WRITER_PROBES: Readonly<Record<string, readonly WriterProbe[]>> = {
  toggleBold: [
    {
      name: "toggle bold",
      slot: "formatted",
      run: (state) => ran(wholeParagraph(state), toggleBold),
    },
  ],
  setTextColor: [
    {
      name: "color the text",
      slot: "formatted",
      run: (state) => ran(wholeParagraph(state), setTextColor(TEXT_COLOR)),
    },
  ],
  setParagraphAlign: [
    {
      name: "align a paragraph",
      slot: "spaced",
      run: (state) =>
        ran(state, setParagraphAlign(otherAlign(activeParagraphAlign(state)))),
    },
  ],
  setLineSpacing: [
    {
      name: "space a paragraph's lines out",
      slot: "spaced",
      run: (state) =>
        ran(state, setLineSpacing(otherSpacing(activeLineSpacing(state)))),
    },
  ],
  toggleNumberedList: [
    {
      name: "start a numbered list",
      slot: "numbered",
      run: (state) => ran(state, toggleNumberedList),
      expect: (exported) => {
        const referenced = referencedNumIds(exported.mainXml);
        expect(
          addedNumIds(exported).filter((numId) => !referenced.has(numId)),
          `${exported.name} list definitions the body points at nowhere`
        ).toEqual([]);
      },
    },
  ],
  toggleBulletList: [
    {
      name: "start a bullet list",
      slot: "bulleted",
      // One definition for each of the two lists the probes above and this one started
      expect: (exported) =>
        expect(
          addedNumIds(exported),
          `${exported.name} new list definitions`
        ).toHaveLength(2),
      run: (state) => ran(state, toggleBulletList),
    },
  ],
  insertTable: [
    {
      name: "insert a table",
      slot: "tabled",
      run: (state) => ran(state, insertTable({ rows: 2, columns: 3 })),
    },
  ],
  addRowAfter: [
    {
      name: "add a row to that table",
      slot: "table",
      run: (state) => ran(state, addRowAfter),
    },
  ],
  mergeCells: [
    {
      name: "merge two cells of the added row",
      slot: "table",
      run: (state) => ran(twoCellsOfRow(state, 1), mergeCells),
    },
  ],
  setCellVerticalAlign: [
    {
      name: "align selected cells vertically",
      slot: "table",
      run: (state) =>
        ran(twoCellsOfRow(state, 1), setCellVerticalAlign("center")),
    },
  ],
  setCellPadding: [
    {
      name: "pad selected cells",
      slot: "table",
      run: (state) =>
        ran(
          twoCellsOfRow(state, 1),
          setCellPadding({ top: 6, right: 8, bottom: 6, left: 8 })
        ),
    },
  ],
  insertImage: [
    {
      name: "insert an image",
      slot: "pictured",
      run: (state) => ran(state, insertImage(A_PICTURE)),
      expect: (exported, untouched) =>
        expectMediaPart(
          exported,
          addedMedia(exported, untouched),
          "png",
          TINY_PNG
        ),
    },
    {
      name: "insert an image of another kind",
      slot: "pictured",
      run: (state) => ran(state, insertImage(ANOTHER_PICTURE)),
      expect: (exported, untouched) => {
        const added = addedMedia(exported, untouched);
        expect(added, `${exported.name} media parts`).toHaveLength(2);
        expectMediaPart(exported, added, "gif", TINY_GIF);

        expect(
          decodeUtf8(openedPart(exported, CONTENT_TYPES_PATH)).text,
          `${exported.name} declares the gif extension already, which leaves the content types writer unreached`
        ).not.toContain('Extension="gif"');
        expect(
          exported.text(CONTENT_TYPES_PATH),
          `${exported.name} content types`
        ).toContain('Extension="gif"');
      },
    },
  ],
  setLink: [
    {
      name: "put a link on a stretch of text",
      slot: "formatted",
      run: (state) => ran(wholeParagraph(state), setLink(LINK_ADDRESS)),
      expect: (exported, untouched) => {
        const links = linkRelationships(exported);
        expect(links.length, `${exported.name} hyperlink relationships`).toBe(
          linkRelationships(untouched).length + 1
        );
        const forTheBattery = links.filter((rel) =>
          rel.includes('Target="https://example.com/battery?a=1&amp;b=2"')
        );
        expect(forTheBattery, `${exported.name} hyperlink target`).toHaveLength(
          1
        );
        expect(
          forTheBattery[0],
          `${exported.name} hyperlink relationship`
        ).toContain('TargetMode="External"');
      },
    },
  ],
  addComment: [
    {
      name: "add a comment to a stretch of text",
      slot: "spaced",
      run: (state) =>
        ran(
          wholeParagraph(state),
          addComment({
            text: "The comment written by the export battery",
            author: "Schema test",
            initials: "ST",
            date: "2026-08-22T00:00:00Z",
          })
        ),
    },
  ],
  // The locks go last: the lock guard turns down every edit reaching into a locked stretch,
  // whichever probe asked for it. The stretch locked first is the one now carrying a link, so the
  // export writes a control around a hyperlink as well
  lockSelection: [
    {
      name: "lock a stretch of text",
      slot: "formatted",
      run: (state) => ran(wholeParagraph(state), lockSelection),
    },
    {
      name: "lock a second stretch of text",
      slot: "spaced",
      run: (state) => ran(wholeParagraph(state), lockSelection),
    },
  ],
  unlockSelection: [
    {
      name: "unlock the second stretch",
      slot: "spaced",
      run: (state) => ran(wholeParagraph(state), unlockSelection),
      // The two locks above minus this one, over the count the same document goes out with
      // untouched, because the caller may have dropped blocks carrying a control of their own
      expect: (exported, untouched) =>
        expect(
          lockedControlCount(exported.mainXml),
          `${exported.name} locked controls`
        ).toBe(lockedControlCount(untouched.mainXml) + 1),
    },
  ],
};

/**
 * Everything else the two entries export, with the reason it reaches no writer.
 *
 * A query answers about the document without changing it, a selection move leaves the document
 * where it was, and undo and redo put back a document a probe above already wrote.
 */
export const NOT_A_WRITER: Readonly<Record<string, string>> = {};

export function everyProbe(): WriterProbe[] {
  return Object.values(WRITER_PROBES).flat();
}

function reasonOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** The state every probe in turn leaves behind */
export function afterTheBattery(state: EditorState): EditorState {
  const places = slotPlaces(state.doc);
  return everyProbe().reduce((before, probe) => {
    try {
      const after = probe.run(placedIn(before, probe.slot, places));
      if (after.doc.eq(before.doc)) {
        throw new Error("the document it left is the one it was given");
      }
      return after;
    } catch (error) {
      throw new Error(
        `the battery could not ${probe.name}: ${reasonOf(error)}`
      );
    }
  }, state);
}

/** What every probe says of the package the export wrote once they had all run */
export function expectProbesWrote(
  exported: ExportedPackage,
  untouched: ExportedPackage
): void {
  for (const probe of everyProbe()) probe.expect?.(exported, untouched);
}
