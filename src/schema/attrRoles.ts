/**
 * What each attr of the schema is for, so that a judgement about a node can say which of them it
 * is judging.
 *
 * An attr is one of three things. A `source` attr is what the exporter writes from: the XML the
 * file arrived as, or a value the model owns that replaces part of it. A `display` attr is worked
 * out from the source and the formatting around it, is never written, and is free to change the
 * moment the surroundings do. A `session` attr says which block, control or link of the open
 * document this is, and is never written either.
 *
 * The split exists because a display value changing is not the document changing. Opening a file
 * in the editor re-derives a table's cell borders (`table/gridBorders`), and a comparison that
 * counted that as an edit would rewrite a table nobody touched, losing the markup the writer does
 * not model. `./sourceEquality` leaves display attrs out for that reason, and keeps session attrs
 * in: two blocks that came from different places in the file are different blocks.
 *
 * `attrRoles.test.ts` holds this table against the schema in both directions, so an attr added
 * without a role fails there before it can reach a judgement that does not know what to do with it.
 */

import { type MarkType, NodeType } from "prosemirror-model";

export type AttrRole = "source" | "display" | "session";

type AttrRoles = Readonly<Record<string, Readonly<Record<string, AttrRole>>>>;

export const NODE_ATTR_ROLES: AttrRoles = {
  paragraph: {
    srcId: "session",
    pAttrs: "source",
    pPr: "source",
    format: "display",
    styleRun: "display",
  },
  table: {
    srcId: "session",
    tblAttrs: "source",
    tblPr: "source",
    tblW: "source",
    gridCols: "source",
    format: "display",
    styleInside: "display",
    styleCellMargins: "display",
  },
  tableRow: {
    trAttrs: "source",
    tblPrEx: "source",
    trPr: "source",
    format: "display",
  },
  tableCell: {
    colspan: "source",
    rowspan: "source",
    // prosemirror-tables' own attr. Import writes null, the writer never reads it, and the table
    // commands only carry it from one cell to another; a real column resize moves `gridCols`
    colwidth: "display",
    tcAttrs: "source",
    tcPr: "source",
    tcW: "source",
    format: "display",
    // Read from `sdtPrefix` rather than from the file, but a lock is not a display value: leaving
    // it out of the comparison would let a step that unlocks a cell pass as a re-derivation
    sdtPrefix: "source",
    sdtContentsLocked: "source",
    sdtDeletionLocked: "source",
  },
  rawBlock: { xml: "source", name: "source" },
  docxRaw: { srcId: "session", name: "session" },
  bookmarkBlock: { srcId: "session", name: "session" },
  hardBreak: { brAttrs: "source" },
  image: { src: "source", extent: "source", alt: "source", xml: "source" },
  commentStart: { id: "source", xml: "source" },
  commentEnd: { id: "source", xml: "source" },
  commentReference: {
    id: "source",
    referenceXml: "source",
    author: "source",
    authorId: "source",
    initials: "source",
    date: "source",
    text: "source",
    commentXml: "source",
    paraId: "source",
    resolved: "source",
    extensionXml: "source",
    replies: "source",
    // Read to decide whether the comment parts are rewritten at all; neither is written
    imported: "session",
    threadImported: "session",
  },
  noteReference: {
    kind: "source",
    id: "source",
    customMarkFollows: "source",
    referenceXml: "source",
    // Both come from the notes part as the document was opened, and the writer puts back the
    // reference alone: a note renumbered around an edit is the same reference it was
    label: "display",
    text: "display",
  },
  rawInline: { xml: "source" },
};

export const MARK_ATTR_ROLES: AttrRoles = {
  run: { rPr: "source", rAttrs: "source", format: "display" },
  sdt: {
    sdtPrefix: "source",
    // Counted through the document as it was opened, to tell one control from the next
    sdtKey: "session",
    contentsLocked: "source",
    deletionLocked: "source",
  },
  link: { linkPrefix: "source", href: "source", linkKey: "session" },
  tab: { tabAttrs: "source" },
};

function rolesOf(
  type: NodeType | MarkType
): Readonly<Record<string, AttrRole>> {
  const roles =
    type instanceof NodeType
      ? NODE_ATTR_ROLES[type.name]
      : MARK_ATTR_ROLES[type.name];
  return roles ?? {};
}

/**
 * What this attr of this node or mark is for.
 *
 * An attr with no role declared is read as `source`: a judgement then compares it, which is the
 * answer that reports a change rather than hiding one. `attrRoles.test.ts` is what keeps the case
 * from arising.
 */
export function attrRole(type: NodeType | MarkType, name: string): AttrRole {
  return rolesOf(type)[name] ?? "source";
}

/** The attrs of this node or mark that are worked out rather than written */
export function displayAttrsOf(type: NodeType | MarkType): readonly string[] {
  return Object.entries(rolesOf(type))
    .filter(([, role]) => role === "display")
    .map(([name]) => name);
}
