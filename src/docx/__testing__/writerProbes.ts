/**
 * The battery of edits that runs before an export is validated, as a probe per public writer.
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
  activeParagraphStyle,
  addComment,
  addCommentReply,
  type DocumentComment,
  type DocumentCommentReply,
  decreaseIndent,
  decreaseListLevel,
  documentComments,
  documentParagraphStyles,
  type ImageToInsert,
  increaseIndent,
  increaseListLevel,
  insertImage,
  insertLineBreak,
  insertPageBreak,
  insertTab,
  insertTable,
  isBoldActive,
  isItalicActive,
  isStrikeActive,
  isUnderlineActive,
  lockSelection,
  type NewComment,
  removeComment,
  removeCommentReply,
  removeLink,
  selectionLock,
  setCommentResolved,
  setFontFamily,
  setFontSize,
  setLineSpacing,
  setLink,
  setParagraphAlign,
  setParagraphStyle,
  setTextBackground,
  setTextColor,
  toggleBold,
  toggleBulletList,
  toggleItalic,
  toggleNumberedList,
  toggleStrike,
  toggleUnderline,
  unlockSelection,
  updateComment,
  updateCommentReply,
} from "../../editor/commands";
import { editorStateForSession } from "../../editor/createEditor";
import {
  type LineSpacing,
  type ParagraphAlign,
  toParagraphFormat,
} from "../../model/format";
import { ST_OnOff } from "../../ooxml/simpleTypes";
import { wAttr } from "../../ooxml/units";
import {
  childByLocalName,
  decodeUtf8,
  namespaceDecls,
  parseXml,
  W_NS,
} from "../../ooxml/xml";
import { sameSource } from "../../schema/sourceEquality";
import {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  deleteColumn,
  deleteRow,
  deleteTable,
  mergeCells,
  setCellBackground,
  setCellBorderColor,
  setCellBorders,
  setCellPadding,
  setCellVerticalAlign,
  splitCell,
} from "../../table";
import { exportDocx } from "../exportDocx";
import type { SessionStore } from "../session";

/**
 * The editing state a screen would hold for this document, opened as the author the comment
 * probes write under. A comment carrying an identity may only be edited by the author it names
 * (`schema/protection`), and the probes that reply to, rewrite and resolve one need that.
 */
export function openState(doc: PMNode, session: SessionStore): EditorState {
  return editorStateForSession(
    { doc, session },
    {
      author: {
        id: AUTHOR.authorId,
        name: AUTHOR.author,
        initials: AUTHOR.initials,
      },
    }
  );
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
const TEXT_BACKGROUND = "#FFF2CC";
const FONT_FAMILY = "Georgia";
const FONT_SIZE_PT = 13;
const CELL_BACKGROUND = "#EAF1F8";
const CELL_BORDER_COLOR = "#C00000";

const COMMENTS_PATH = "word/comments.xml";
const COMMENTS_EXTENDED_PATH = "word/commentsExtended.xml";

/**
 * The author of everything the comment probes write. The identity is what the export records in
 * the people part, which is written for no other reason and would go unreached without it
 */
const AUTHOR = {
  author: "Schema test",
  authorId: "schema-test",
  initials: "ST",
};
const WRITTEN_AT = "2026-08-22T00:00:00Z";

const A_COMMENT: NewComment = {
  ...AUTHOR,
  date: WRITTEN_AT,
  text: "The comment written by the export battery",
};

/** The comment a probe below takes away again */
const ANOTHER_COMMENT: NewComment = {
  ...AUTHOR,
  date: WRITTEN_AT,
  text: "The comment the battery takes away again",
};

const EDITED_COMMENT = "The comment the battery rewrote";

const A_REPLY: NewComment = {
  ...AUTHOR,
  date: WRITTEN_AT,
  text: "The reply written by the export battery",
};

/** The reply a probe below takes away again */
const ANOTHER_REPLY: NewComment = {
  ...AUTHOR,
  date: WRITTEN_AT,
  text: "The reply the battery takes away again",
};

const EDITED_REPLY = "The reply the battery rewrote";

/** A paragraph style the document defines that the paragraph is not already written in */
function otherStyleId(state: EditorState): string {
  const active = activeParagraphStyle(state);
  const worn = active.kind === "shared" ? active.styleId : null;
  const option = documentParagraphStyles(state).find(
    (style) => !style.isDefault && !style.hidden && style.id !== worn
  );
  if (option === undefined) {
    throw new Error("the document defines no other paragraph style to put on");
  }
  return option.id;
}

/** The comment a probe left behind, which the ones working on it find it by */
function commentReading(state: EditorState, text: string): DocumentComment {
  const comment = documentComments(state).find((entry) => entry.text === text);
  if (comment === undefined) {
    throw new Error(`the document holds no comment reading "${text}"`);
  }
  return comment;
}

function replyReading(
  comment: DocumentComment,
  text: string
): DocumentCommentReply {
  const reply = comment.replies.find((entry) => entry.text === text);
  if (reply === undefined) {
    throw new Error(`the comment holds no reply reading "${text}"`);
  }
  return reply;
}

/** The address a probe links a stretch of text to, which the export writes a relationship for */
const LINK_ADDRESS = "https://example.com/battery?a=1&b=2";

/** The address of the link a probe puts on and the one after it takes off again */
const REMOVED_ADDRESS = "https://example.com/battery/taken-off";

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

/** Each probe must check its immediate result; package-wide checks run after the whole battery. */
export interface WriterProbe {
  /** What the probe does, which is what a battery that could not run it is reported by */
  name: string;
  slot: ProbeSlot;
  prepare?(state: EditorState): EditorState;
  run(state: EditorState): EditorState;
  check(before: EditorState, after: EditorState): void;
  expect?(exported: ExportedPackage, untouched: ExportedPackage): void;
}

type ProbeCheck = (before: EditorState, after: EditorState) => void;

function property(xml: unknown, tag: string, attribute = "val"): string | null {
  if (typeof xml !== "string") return null;
  let node: Element | undefined = parseXml(
    `<x ${namespaceDecls(xml)}>${xml}</x>`
  ).documentElement;
  for (const name of tag.split("/"))
    node = node?.getElementsByTagNameNS(W_NS, name)[0];
  return node?.getAttributeNS(W_NS, attribute) ?? null;
}

function runProperty(
  tag: string,
  value: string,
  attribute = "val"
): ProbeCheck {
  return (_before, after) => {
    const runs: PMNode[] = [];
    after.selection.$from.parent.descendants((node) => {
      if (node.isText) runs.push(node);
    });
    expect(runs.length).toBeGreaterThan(0);
    for (const node of runs) {
      const run = node.marks.find((mark) => mark.type.name === "run");
      expect(property(run?.attrs.rPr, tag, attribute)).toBe(value);
    }
  };
}

function toggleProperty(
  tag: string,
  active: (state: EditorState) => boolean
): ProbeCheck {
  return (before, after) => {
    const on = !active(wholeParagraph(before));
    const values: boolean[] = [];
    after.selection.$from.parent.descendants((node) => {
      if (!node.isText) return;
      const xml: unknown = node.marks.find((mark) => mark.type.name === "run")
        ?.attrs.rPr;
      const el =
        typeof xml === "string"
          ? parseXml(
              `<x ${namespaceDecls(xml)}>${xml}</x>`
            ).getElementsByTagNameNS(W_NS, tag)[0]
          : undefined;
      const written = el?.getAttributeNS(W_NS, "val") ?? null;
      // `w:u w:val="none"` is an underline of no kind, which this probe counts as off as well
      values.push(
        el !== undefined &&
          written !== "none" &&
          (ST_OnOff.parse(written) ?? true)
      );
    });
    expect(values.length).toBeGreaterThan(0);
    expect(values.every((value) => value === on)).toBe(true);
  };
}

function paragraphProperty(
  state: EditorState,
  tag: string,
  attribute = "val"
): string | null {
  return property(state.selection.$from.parent.attrs.pPr, tag, attribute);
}

function indentCheck(delta: number): ProbeCheck {
  return (before, after) => {
    const format = toParagraphFormat(
      before.selection.$from.parent.attrs.format
    );
    const left = (format?.indentStartPt ?? format?.indentLeftPt ?? 0) * 20;
    expect(
      Number(
        paragraphProperty(after, "ind", "left") ??
          paragraphProperty(after, "ind", "start") ??
          0
      )
    ).toBe(Math.max(0, Math.round(left) + delta));
  };
}

function levelCheck(delta: number): ProbeCheck {
  return (before, after) => {
    expect(Number(paragraphProperty(after, "ilvl"))).toBe(
      Number(paragraphProperty(before, "ilvl")) + delta
    );
  };
}

function nodeCount(doc: PMNode, name: string): number {
  let count = 0;
  doc.descendants((node) => {
    if (node.type.name === name) count += 1;
  });
  return count;
}

function nodeAdded(name: string, delta = 1): ProbeCheck {
  return (before, after) =>
    expect(nodeCount(after.doc, name)).toBe(
      nodeCount(before.doc, name) + delta
    );
}

function breakAdded(page: boolean): ProbeCheck {
  return (before, after) => {
    const count = (doc: PMNode): number => {
      let found = 0;
      doc.descendants((node) => {
        if (node.type.name !== "hardBreak") return;
        const attrs: unknown = node.attrs.brAttrs;
        const xml = typeof attrs === "string" ? `<w:br ${attrs}/>` : "<w:br/>";
        if ((property(xml, "br", "type") === "page") === page) found += 1;
      });
      return found;
    };
    expect(count(after.doc)).toBe(count(before.doc) + 1);
  };
}

function tableSizeChange(width: number, height: number): ProbeCheck {
  return (before, after) => {
    const a = TableMap.get(tableAround(before).node);
    const b = TableMap.get(tableAround(after).node);
    expect([b.width, b.height]).toEqual([a.width + width, a.height + height]);
  };
}

function unmergedColumn(state: EditorState): EditorState {
  const table = tableAround(state);
  const map = TableMap.get(table.node);
  // After deleting the first row, the new first row starts with a two-column merged cell.
  // Select the split row below it so this probe deletes exactly one column.
  return caretAt(state, table.pos + 1 + map.positionAt(1, 0, table.node) + 1);
}

function cellProperty(
  tag: string,
  value: string,
  attribute = "val"
): ProbeCheck {
  return (_before, after) => {
    expect(after.selection).toBeInstanceOf(CellSelection);
    if (!(after.selection instanceof CellSelection))
      throw new Error("No cells selected");
    after.selection.forEachCell((cell) =>
      expect(property(cell.attrs.tcPr, tag, attribute)).toBe(value)
    );
  };
}

function markCount(doc: PMNode, name: string): number {
  let count = 0;
  doc.descendants((node) => {
    count += node.marks.filter((mark) => mark.type.name === name).length;
  });
  return count;
}

function addedComment(text: string): ProbeCheck {
  return (before, after) => {
    expect(documentComments(after).length).toBe(
      documentComments(before).length + 1
    );
    expect(
      documentComments(after).some((comment) => comment.text === text)
    ).toBe(true);
  };
}

function addedReply(text: string): ProbeCheck {
  return (before, after) => {
    const a = commentReading(before, A_COMMENT.text);
    const b = commentReading(after, A_COMMENT.text);
    expect(b.replies.length).toBe(a.replies.length + 1);
    expect(b.replies.some((reply) => reply.text === text)).toBe(true);
  };
}

function linkedTo(address: string): ProbeCheck {
  return (_before, after) => {
    const marks = after.selection.$from.parent.firstChild?.marks ?? [];
    expect(marks.find((mark) => mark.type.name === "link")?.attrs.href).toBe(
      address
    );
  };
}

export const WRITER_PROBES: Readonly<Record<string, readonly WriterProbe[]>> = {
  toggleBold: [
    {
      name: "toggle bold",
      slot: "formatted",
      check: toggleProperty("b", isBoldActive),
      run: (state) => ran(wholeParagraph(state), toggleBold),
    },
  ],
  toggleItalic: [
    {
      name: "toggle italic",
      slot: "formatted",
      check: toggleProperty("i", isItalicActive),
      run: (state) => ran(wholeParagraph(state), toggleItalic),
    },
  ],
  toggleUnderline: [
    {
      name: "toggle underline",
      slot: "formatted",
      check: toggleProperty("u", isUnderlineActive),
      run: (state) => ran(wholeParagraph(state), toggleUnderline),
    },
  ],
  toggleStrike: [
    {
      name: "strike the text through",
      slot: "formatted",
      check: toggleProperty("strike", isStrikeActive),
      run: (state) => ran(wholeParagraph(state), toggleStrike),
    },
  ],
  setTextColor: [
    {
      name: "color the text",
      slot: "formatted",
      check: runProperty("color", TEXT_COLOR.slice(1)),
      run: (state) => ran(wholeParagraph(state), setTextColor(TEXT_COLOR)),
    },
  ],
  setTextBackground: [
    {
      name: "highlight the text",
      slot: "formatted",
      check: runProperty("shd", TEXT_BACKGROUND.slice(1), "fill"),
      run: (state) =>
        ran(wholeParagraph(state), setTextBackground(TEXT_BACKGROUND)),
    },
  ],
  setFontFamily: [
    {
      name: "set the font",
      slot: "formatted",
      check: runProperty("rFonts", FONT_FAMILY, "ascii"),
      run: (state) => ran(wholeParagraph(state), setFontFamily(FONT_FAMILY)),
    },
  ],
  setFontSize: [
    {
      name: "set the font size",
      slot: "formatted",
      check: runProperty("sz", String(FONT_SIZE_PT * 2)),
      run: (state) => ran(wholeParagraph(state), setFontSize(FONT_SIZE_PT)),
    },
  ],
  setParagraphStyle: [
    {
      name: "put a paragraph style on a paragraph",
      slot: "formatted",
      check: (before, after) =>
        expect(paragraphProperty(after, "pStyle")).toBe(otherStyleId(before)),
      run: (state) => ran(state, setParagraphStyle(otherStyleId(state))),
    },
  ],
  setParagraphAlign: [
    {
      name: "align a paragraph",
      slot: "spaced",
      check: (before, after) => {
        const align = otherAlign(activeParagraphAlign(before));
        expect(paragraphProperty(after, "jc")).toBe(
          align === "justify" ? "both" : align
        );
      },
      run: (state) =>
        ran(state, setParagraphAlign(otherAlign(activeParagraphAlign(state)))),
    },
  ],
  setLineSpacing: [
    {
      name: "space a paragraph's lines out",
      slot: "spaced",
      check: (before, after) => {
        const spacing = otherSpacing(activeLineSpacing(before));
        expect(paragraphProperty(after, "spacing", "line")).toBe(
          String(spacing.rule === "auto" ? spacing.lines * 240 : 0)
        );
      },
      run: (state) =>
        ran(state, setLineSpacing(otherSpacing(activeLineSpacing(state)))),
    },
  ],
  increaseIndent: [
    {
      name: "indent a paragraph",
      slot: "spaced",
      check: indentCheck(720),
      run: (state) => ran(state, increaseIndent),
    },
  ],
  decreaseIndent: [
    {
      name: "take that indent back off",
      slot: "spaced",
      check: indentCheck(-720),
      run: (state) => ran(state, decreaseIndent),
    },
  ],
  toggleNumberedList: [
    {
      name: "start a numbered list",
      slot: "numbered",
      check: (before, after) => {
        expect(paragraphProperty(after, "numId")).not.toBeNull();
        expect(paragraphProperty(after, "numId")).not.toBe(
          paragraphProperty(before, "numId")
        );
      },
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
  increaseListLevel: [
    {
      name: "move that list item a level deeper",
      slot: "numbered",
      check: levelCheck(1),
      run: (state) => ran(state, increaseListLevel),
    },
  ],
  decreaseListLevel: [
    {
      name: "move that list item a level back up",
      slot: "numbered",
      check: levelCheck(-1),
      run: (state) => ran(state, decreaseListLevel),
    },
  ],
  toggleBulletList: [
    {
      name: "start a bullet list",
      slot: "bulleted",
      check: (before, after) => {
        expect(paragraphProperty(after, "numId")).not.toBeNull();
        expect(paragraphProperty(after, "numId")).not.toBe(
          paragraphProperty(before, "numId")
        );
      },
      // One definition for each of the two lists this probe and the numbered one above started
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
      check: (before, after) => {
        nodeAdded("table")(before, after);
        const map = TableMap.get(tableAround(after).node);
        expect([map.width, map.height]).toEqual([4, 3]);
      },
      run: (state) => ran(state, insertTable({ rows: 3, columns: 4 })),
    },
  ],
  addRowAfter: [
    {
      name: "add a row under the first",
      slot: "table",
      check: tableSizeChange(0, 1),
      run: (state) => ran(state, addRowAfter),
    },
  ],
  addRowBefore: [
    {
      name: "add a row over the first",
      slot: "table",
      check: tableSizeChange(0, 1),
      run: (state) => ran(state, addRowBefore),
    },
  ],
  addColumnAfter: [
    {
      name: "add a column beside the first",
      slot: "table",
      check: tableSizeChange(1, 0),
      run: (state) => ran(state, addColumnAfter),
    },
  ],
  addColumnBefore: [
    {
      name: "add a column before the first",
      slot: "table",
      check: tableSizeChange(1, 0),
      run: (state) => ran(state, addColumnBefore),
    },
  ],
  mergeCells: [
    {
      name: "merge two cells of the second row",
      slot: "table",
      check: nodeAdded("tableCell", -1),
      run: (state) => ran(twoCellsOfRow(state, 1), mergeCells),
    },
    // The cell the split below takes apart. Merging one row and splitting another is what leaves
    // the package with both a merged cell and a cell the writer built out of one
    {
      name: "merge two cells of the third row",
      slot: "table",
      check: nodeAdded("tableCell", -1),
      run: (state) => ran(twoCellsOfRow(state, 2), mergeCells),
    },
  ],
  splitCell: [
    {
      name: "split the merged cell of the third row",
      slot: "table",
      check: nodeAdded("tableCell"),
      run: (state) => ran(twoCellsOfRow(state, 2), splitCell),
    },
  ],
  setCellBackground: [
    {
      name: "shade selected cells",
      slot: "table",
      check: cellProperty("shd", CELL_BACKGROUND.slice(1), "fill"),
      run: (state) =>
        ran(twoCellsOfRow(state, 3), setCellBackground(CELL_BACKGROUND)),
    },
  ],
  setCellBorderColor: [
    {
      name: "color the borders of selected cells",
      slot: "table",
      check: cellProperty("top", CELL_BORDER_COLOR.slice(1), "color"),
      run: (state) =>
        ran(twoCellsOfRow(state, 3), setCellBorderColor(CELL_BORDER_COLOR)),
    },
  ],
  setCellBorders: [
    {
      // A row of its own, and the preset that clears the lines: `outer` draws a single line on
      // the sides of the selection that face outward, which the cells of a new table are already
      // drawn with, so it would report that it changes nothing
      name: "clear the lines of the cells of a row",
      slot: "table",
      check: cellProperty("top", "none"),
      run: (state) => ran(twoCellsOfRow(state, 4), setCellBorders("none")),
    },
  ],
  setCellVerticalAlign: [
    {
      name: "align selected cells vertically",
      slot: "table",
      check: cellProperty("vAlign", "center"),
      run: (state) =>
        ran(twoCellsOfRow(state, 3), setCellVerticalAlign("center")),
    },
  ],
  setCellPadding: [
    {
      name: "pad selected cells",
      slot: "table",
      check: cellProperty("tcMar/top", "120", "w"),
      run: (state) =>
        ran(
          twoCellsOfRow(state, 3),
          setCellPadding({ top: 6, right: 8, bottom: 6, left: 8 })
        ),
    },
  ],
  deleteRow: [
    {
      name: "delete the first row of that table",
      slot: "table",
      check: tableSizeChange(0, -1),
      run: (state) => ran(state, deleteRow),
    },
  ],
  deleteColumn: [
    {
      name: "delete the first column of that table",
      slot: "table",
      check: tableSizeChange(-1, 0),
      run: (state) => ran(unmergedColumn(state), deleteColumn),
    },
  ],
  deleteTable: [
    {
      name: "delete a table",
      slot: "tabled",
      check: nodeAdded("table", -1),
      // The table this one takes away is one it puts there itself, so that the table every probe
      // above worked in is the one written out. What stays behind is the empty paragraph a new
      // table is inserted with
      prepare: (state) => ran(state, insertTable({ rows: 2, columns: 2 })),
      run: (state) => ran(state, deleteTable),
    },
  ],
  insertImage: [
    {
      name: "insert an image",
      slot: "pictured",
      check: nodeAdded("image"),
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
      check: nodeAdded("image"),
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
  insertLineBreak: [
    {
      name: "break a line",
      slot: "pictured",
      check: breakAdded(false),
      run: (state) => ran(state, insertLineBreak),
    },
  ],
  insertPageBreak: [
    {
      name: "break a page",
      slot: "pictured",
      check: breakAdded(true),
      run: (state) => ran(state, insertPageBreak),
    },
  ],
  insertTab: [
    {
      name: "put a tab in",
      slot: "pictured",
      check: (before, after) =>
        expect(markCount(after.doc, "tab")).toBe(
          markCount(before.doc, "tab") + 1
        ),
      run: (state) => ran(state, insertTab),
    },
  ],
  setLink: [
    {
      name: "put a link on a stretch of text",
      slot: "formatted",
      check: linkedTo(LINK_ADDRESS),
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
    // The link the probe below takes off again, which is why the relationship count above counts
    // one link and not two
    {
      name: "put a second link on a stretch of text",
      slot: "bulleted",
      check: linkedTo(REMOVED_ADDRESS),
      run: (state) => ran(wholeParagraph(state), setLink(REMOVED_ADDRESS)),
    },
  ],
  removeLink: [
    {
      name: "take that second link off again",
      slot: "bulleted",
      check: (_before, after) =>
        expect(markCount(after.selection.$from.parent, "link")).toBe(0),
      run: (state) => ran(wholeParagraph(state), removeLink),
      expect: (exported) =>
        expect(
          linkRelationships(exported).filter((rel) =>
            rel.includes(REMOVED_ADDRESS)
          ),
          `${exported.name} relationships left behind by a link that was taken off`
        ).toEqual([]),
    },
  ],
  addComment: [
    {
      name: "add a comment to a stretch of text",
      slot: "spaced",
      check: addedComment(A_COMMENT.text),
      run: (state) => ran(wholeParagraph(state), addComment(A_COMMENT)),
    },
    // The comment the probe below takes away again
    {
      name: "add a second comment",
      slot: "bulleted",
      check: addedComment(ANOTHER_COMMENT.text),
      run: (state) => ran(wholeParagraph(state), addComment(ANOTHER_COMMENT)),
    },
  ],
  addCommentReply: [
    {
      name: "reply to that comment",
      slot: "spaced",
      check: addedReply(A_REPLY.text),
      run: (state) =>
        ran(
          state,
          addCommentReply(commentReading(state, A_COMMENT.text).id, A_REPLY)
        ),
    },
    // The reply the removal below takes away again
    {
      name: "reply to it a second time",
      slot: "spaced",
      check: addedReply(ANOTHER_REPLY.text),
      run: (state) =>
        ran(
          state,
          addCommentReply(
            commentReading(state, A_COMMENT.text).id,
            ANOTHER_REPLY
          )
        ),
    },
  ],
  updateCommentReply: [
    {
      name: "rewrite the first reply",
      slot: "spaced",
      check: (_before, after) =>
        expect(
          commentReading(after, A_COMMENT.text).replies.some(
            (reply) => reply.text === EDITED_REPLY
          )
        ).toBe(true),
      run: (state) => {
        const comment = commentReading(state, A_COMMENT.text);
        return ran(
          state,
          updateCommentReply(
            comment.id,
            replyReading(comment, A_REPLY.text).id,
            EDITED_REPLY
          )
        );
      },
    },
  ],
  removeCommentReply: [
    {
      name: "take the second reply away",
      slot: "spaced",
      check: (before, after) => {
        const a = commentReading(before, A_COMMENT.text);
        const b = commentReading(after, A_COMMENT.text);
        expect(b.replies.length).toBe(a.replies.length - 1);
        expect(
          b.replies.some((reply) => reply.text === ANOTHER_REPLY.text)
        ).toBe(false);
      },
      run: (state) => {
        const comment = commentReading(state, A_COMMENT.text);
        return ran(
          state,
          removeCommentReply(
            comment.id,
            replyReading(comment, ANOTHER_REPLY.text).id
          )
        );
      },
      expect: (exported) =>
        expect(
          exported.text(COMMENTS_PATH),
          `${exported.name} keeps a reply that was taken away`
        ).not.toContain(ANOTHER_REPLY.text),
    },
  ],
  updateComment: [
    {
      name: "rewrite that comment",
      slot: "spaced",
      check: (_before, after) =>
        expect(commentReading(after, EDITED_COMMENT).text).toBe(EDITED_COMMENT),
      run: (state) =>
        ran(
          state,
          updateComment(
            commentReading(state, A_COMMENT.text).id,
            EDITED_COMMENT
          )
        ),
      expect: (exported) => {
        const comments = exported.text(COMMENTS_PATH);
        expect(comments, `${exported.name} comments part`).toContain(
          EDITED_COMMENT
        );
        expect(
          comments,
          `${exported.name} keeps the text a comment was rewritten from`
        ).not.toContain(A_COMMENT.text);
      },
    },
  ],
  setCommentResolved: [
    {
      name: "resolve that thread",
      slot: "spaced",
      check: (_before, after) =>
        expect(commentReading(after, EDITED_COMMENT).resolved).toBe(true),
      run: (state) =>
        ran(
          state,
          setCommentResolved(commentReading(state, EDITED_COMMENT).id, true)
        ),
      // The thread state is what the part beside the comments part carries, and the key it hangs
      // off is `w14:paraId` on the comment's own paragraph, which is markup the part 1 schemas
      // describe nowhere. Reading it takes the preprocessing this suite validates through
      expect: (exported) =>
        expect(
          exported.text(COMMENTS_EXTENDED_PATH),
          `${exported.name} thread state`
        ).toContain('w15:done="1"'),
    },
  ],
  removeComment: [
    {
      name: "take the second comment away",
      slot: "bulleted",
      check: (before, after) => {
        expect(documentComments(after).length).toBe(
          documentComments(before).length - 1
        );
        expect(
          documentComments(after).some(
            (comment) => comment.text === ANOTHER_COMMENT.text
          )
        ).toBe(false);
      },
      run: (state) =>
        ran(
          state,
          removeComment(commentReading(state, ANOTHER_COMMENT.text).id)
        ),
      expect: (exported) =>
        expect(
          exported.text(COMMENTS_PATH),
          `${exported.name} keeps a comment that was taken away`
        ).not.toContain(ANOTHER_COMMENT.text),
    },
  ],
  // The locks go last: the lock guard turns down every edit reaching into a locked stretch,
  // whichever probe asked for it. The stretch locked first is the one now carrying a link, so the
  // export writes a control around a hyperlink as well
  lockSelection: [
    {
      name: "lock a stretch of text",
      slot: "formatted",
      check: (_before, after) =>
        expect(selectionLock(wholeParagraph(after))).toBe("locked"),
      run: (state) => ran(wholeParagraph(state), lockSelection),
    },
    {
      name: "lock a second stretch of text",
      slot: "spaced",
      check: (_before, after) =>
        expect(selectionLock(wholeParagraph(after))).toBe("locked"),
      run: (state) => ran(wholeParagraph(state), lockSelection),
    },
  ],
  unlockSelection: [
    {
      name: "unlock the second stretch",
      slot: "spaced",
      check: (_before, after) =>
        expect(selectionLock(wholeParagraph(after))).toBe("lockable"),
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
 * A query answers about the document and leaves it as it was, and so does a selection move. Undo
 * and redo put back a document a probe above already wrote, so what they reach the writer with is
 * markup another probe already had it write.
 */
export const NOT_A_WRITER: Readonly<Record<string, string>> = {
  IMAGE_FILE_ACCEPT: "the accept string a file picker is given",
  SINGLE_LINE_SPACING: "a line spacing value",
  activeCellBackground: "a query about the selected cells",
  activeCellBorderColor: "a query about the selected cells",
  activeCellPadding: "a query about the selected cells",
  activeCellVerticalAlign: "a query about the selected cells",
  activeFontFamily: "a query about the selection",
  activeFontSize: "a query about the selection",
  activeLineSpacing: "a query about the selection",
  activeLink: "a query about the selection",
  activeLinkSpan: "a query about the selection",
  activeListKind: "a query about the selection",
  activeParagraphAlign: "a query about the selection",
  activeParagraphStyle: "a query about the selection",
  activeTextBackground: "a query about the selection",
  activeTextColor: "a query about the selection",
  canAddComment: "the query the add-comment button is drawn from",
  canDecreaseIndent: "the query the decrease-indent button is drawn from",
  canEditComment:
    "the query the edit and delete buttons of a comment are drawn from",
  canExport: "the query an export control is drawn from",
  canFormatText: "the query the character formatting controls are drawn from",
  canIncreaseIndent: "the query the increase-indent button is drawn from",
  canInsertImage: "the query the image button is drawn from",
  canInsertTable: "the query the insert-table button is drawn from",
  canMergeCells: "the query the merge row is drawn from",
  canRunCommand: "the question asked of a command the package does not own",
  canSetCellBorderColor: "the query the border colour picker is drawn from",
  canSetCellFormatting: "the query the cell layout controls are drawn from",
  canSetLineSpacing: "the query the spacing menu is drawn from",
  canSetLink: "the query the link button is drawn from",
  canSetParagraphAlign: "the query the alignment menu is drawn from",
  canSplitCell: "the query the split row is drawn from",
  documentBodyWidthPx: "a measurement read off the open document",
  documentComments: "the comments displayed alongside the document",
  documentDefaults: "the formatting the document declares",
  documentExportProblems: "a query about the document, writing nothing",
  documentFidelity: "a query about the document, writing nothing",
  documentFontNames: "the fonts the document names",
  documentHasLocked: "a query about the document",
  documentNotes: "the notes displayed after the document",
  documentParagraphStyles: "the styles the document defines",
  editingProtection: "a query about what the editor as a whole may receive",
  fittedExtent: "the rule an oversized image is shrunk by",
  imageFilesIn: "picks the image files out of a picker, clipboard or drag",
  insertImageFiles:
    "reads the files first, then runs `insertImage`, whose probe is above",
  isBoldActive: "a query about the selection",
  isInList: "a query about the selection",
  isInTable: "the query the table buttons are drawn from",
  isItalicActive: "a query about the selection",
  isStrikeActive: "a query about the selection",
  isUnderlineActive: "a query about the selection",
  readImageFile: "reads one file and gives the size it comes in at",
  redo: "puts back a document the probe it replays already wrote",
  selectComment: "moves the selection to a comment's anchor",
  selectionLock: "a query about the selection",
  selectionTouchesLocked: "a query about the selection",
  undo: "puts back a document the probe before it already wrote",
};

/** Every probe of the registry, in the order it is written in, which is the order they run in */
function everyProbe(): WriterProbe[] {
  return Object.values(WRITER_PROBES).flat();
}

function reasonOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** The state every probe in turn leaves behind */
export function afterTheBattery(
  state: EditorState,
  inspect?: (
    probe: WriterProbe,
    before: EditorState,
    after: EditorState
  ) => void
): EditorState {
  const places = slotPlaces(state.doc);
  return everyProbe().reduce((before, probe) => {
    try {
      const placed = placedIn(before, probe.slot, places);
      const prepared = probe.prepare?.(placed) ?? placed;
      const after = probe.run(prepared);
      if (sameSource(after.doc, prepared.doc)) {
        throw new Error("the document it left is the one it was given");
      }
      probe.check(prepared, after);
      inspect?.(probe, prepared, after);
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
