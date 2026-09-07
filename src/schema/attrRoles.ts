/**
 * What each attr of the schema is for, so that a judgement about a node can say which of them it
 * is judging, and so that a consumer can tell which of them are theirs to rely on.
 *
 * Two things are said about every attr, because they answer different questions.
 *
 * Its **role** is what a comparison does with it. A `source` attr is what the exporter writes
 * from: the XML the file arrived as, or a value the model owns that replaces part of it. A
 * `display` attr is worked out from the source and the formatting around it, is never written, and
 * is free to change the moment the surroundings do. A `session` attr says which block, control or
 * link of the open document this is, and is never written either.
 *
 * The split exists because a display value changing is not the document changing. Opening a file
 * in the editor re-derives a table's cell borders (`table/gridBorders`), and a comparison that
 * counted that as an edit would rewrite a table nobody touched, losing the markup the writer does
 * not model. `./sourceEquality` leaves display attrs out for that reason, and keeps session attrs
 * in: two blocks that came from different places in the file are different blocks.
 *
 * Its **class** is where the value comes from, which is the line the published boundary is drawn
 * along. A `preserved` attr holds OOXML the file arrived as and export writes that string back as
 * it stands. A `derived` attr is worked out again from another attr or from the formatting around
 * it and never reaches the file. An `identity` attr says where the node came from or what it
 * stands for, and the editor does not make the value up. Everything else is `model`: a value the
 * editor owns and the commands change.
 *
 * The two are not one column under two names. A lock flag is `derived`, since import reads it once
 * out of the control's own XML, but its role is `source`: leaving it out of a comparison would let
 * a step that unlocks a cell pass as a re-derivation. `imported` and `threadImported` are `derived`
 * for the same reason and `session` by role.
 *
 * `attrClasses.test.ts` and `attrRoles.test.ts` hold this table against the schema in both
 * directions, so an attr added without an entry fails there before it can reach a judgement that
 * does not know what to do with it.
 */

import { type MarkType, NodeType } from "prosemirror-model";

export type AttrRole = "source" | "display" | "session";

export type AttrClass = "preserved" | "derived" | "identity" | "model";

/** What is declared about one attr: what a comparison does with it, and where its value comes from */
export type AttrFacts = { readonly role: AttrRole; readonly class: AttrClass };

type AttrTable = Readonly<Record<string, Readonly<Record<string, AttrFacts>>>>;

export const NODE_ATTR_ROLES: AttrTable = {
  paragraph: {
    srcId: { role: "session", class: "identity" },
    pAttrs: { role: "source", class: "preserved" },
    pPr: { role: "source", class: "preserved" },
    format: { role: "display", class: "derived" },
    styleRun: { role: "display", class: "derived" },
  },
  table: {
    srcId: { role: "session", class: "identity" },
    tblAttrs: { role: "source", class: "preserved" },
    tblPr: { role: "source", class: "preserved" },
    tblW: { role: "source", class: "model" },
    gridCols: { role: "source", class: "model" },
    // The grid the table had before it was last revised, carried as it arrived
    gridChange: { role: "source", class: "preserved" },
    format: { role: "display", class: "derived" },
    styleInside: { role: "display", class: "derived" },
    styleCellMargins: { role: "display", class: "derived" },
  },
  tableRow: {
    trAttrs: { role: "source", class: "preserved" },
    tblPrEx: { role: "source", class: "preserved" },
    trPr: { role: "source", class: "preserved" },
    format: { role: "display", class: "derived" },
  },
  tableCell: {
    colspan: { role: "source", class: "model" },
    rowspan: { role: "source", class: "model" },
    // prosemirror-tables' own attr. Import writes null, the writer never reads it, and the table
    // commands only carry it from one cell to another; a real column resize moves `gridCols`
    colwidth: { role: "display", class: "derived" },
    tcAttrs: { role: "source", class: "preserved" },
    tcPr: { role: "source", class: "preserved" },
    tcW: { role: "source", class: "model" },
    format: { role: "display", class: "derived" },
    // Read from `sdtPrefix` rather than from the file, but a lock is not a display value: leaving
    // it out of the comparison would let a step that unlocks a cell pass as a re-derivation
    sdtPrefix: { role: "source", class: "preserved" },
    sdtContentsLocked: { role: "source", class: "derived" },
    sdtDeletionLocked: { role: "source", class: "derived" },
  },
  rawBlock: {
    xml: { role: "source", class: "preserved" },
    name: { role: "source", class: "identity" },
  },
  docxRaw: {
    srcId: { role: "session", class: "identity" },
    name: { role: "session", class: "identity" },
  },
  bookmarkBlock: {
    srcId: { role: "session", class: "identity" },
    name: { role: "session", class: "identity" },
  },
  hardBreak: { brAttrs: { role: "source", class: "preserved" } },
  image: {
    // The bytes themselves, as a data URL the editor makes when an image is inserted
    src: { role: "source", class: "model" },
    extent: { role: "source", class: "model" },
    alt: { role: "source", class: "model" },
    xml: { role: "source", class: "preserved" },
  },
  commentStart: {
    id: { role: "source", class: "identity" },
    xml: { role: "source", class: "preserved" },
  },
  commentEnd: {
    id: { role: "source", class: "identity" },
    xml: { role: "source", class: "preserved" },
  },
  commentReference: {
    id: { role: "source", class: "identity" },
    referenceXml: { role: "source", class: "preserved" },
    author: { role: "source", class: "model" },
    authorId: { role: "source", class: "model" },
    initials: { role: "source", class: "model" },
    date: { role: "source", class: "model" },
    text: { role: "source", class: "model" },
    commentXml: { role: "source", class: "preserved" },
    paraId: { role: "source", class: "identity" },
    resolved: { role: "source", class: "model" },
    extensionXml: { role: "source", class: "preserved" },
    replies: { role: "source", class: "model" },
    // Read to decide whether the comment parts are rewritten at all; neither is written
    imported: { role: "session", class: "derived" },
    threadImported: { role: "session", class: "derived" },
  },
  noteReference: {
    kind: { role: "source", class: "model" },
    id: { role: "source", class: "identity" },
    customMarkFollows: { role: "source", class: "model" },
    referenceXml: { role: "source", class: "preserved" },
    // Both come from the notes part as the document was opened, and the writer puts back the
    // reference alone: a note renumbered around an edit is the same reference it was
    label: { role: "display", class: "derived" },
    text: { role: "display", class: "derived" },
  },
  rawInline: { xml: { role: "source", class: "preserved" } },
};

export const MARK_ATTR_ROLES: AttrTable = {
  run: {
    rPr: { role: "source", class: "preserved" },
    rAttrs: { role: "source", class: "preserved" },
    format: { role: "display", class: "derived" },
  },
  sdt: {
    sdtPrefix: { role: "source", class: "preserved" },
    // Counted through the document as it was opened, to tell one control from the next
    sdtKey: { role: "session", class: "identity" },
    contentsLocked: { role: "source", class: "derived" },
    deletionLocked: { role: "source", class: "derived" },
  },
  link: {
    linkPrefix: { role: "source", class: "preserved" },
    href: { role: "source", class: "model" },
    linkKey: { role: "session", class: "identity" },
  },
  tab: { tabAttrs: { role: "source", class: "preserved" } },
};

function factsOf(
  type: NodeType | MarkType
): Readonly<Record<string, AttrFacts>> {
  const facts =
    type instanceof NodeType
      ? NODE_ATTR_ROLES[type.name]
      : MARK_ATTR_ROLES[type.name];
  return facts ?? {};
}

/**
 * What this attr of this node or mark is for.
 *
 * An attr with no role declared is read as `source`: a judgement then compares it, which is the
 * answer that reports a change rather than hiding one. `attrRoles.test.ts` is what keeps the case
 * from arising.
 */
export function attrRole(type: NodeType | MarkType, name: string): AttrRole {
  return factsOf(type)[name]?.role ?? "source";
}

/** The attrs of this node or mark that are worked out rather than written */
export function displayAttrsOf(type: NodeType | MarkType): readonly string[] {
  return Object.entries(factsOf(type))
    .filter(([, facts]) => facts.role === "display")
    .map(([name]) => name);
}

/** The attrs of this node or mark that the given class covers */
export function attrsOfClass(
  type: NodeType | MarkType,
  cls: AttrClass
): readonly string[] {
  return Object.entries(factsOf(type))
    .filter(([, facts]) => facts.class === cls)
    .map(([name]) => name);
}
