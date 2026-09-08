/**
 * Each attr has two independent classifications. Its role controls source comparison: `source`
 * values are what the writer writes a block from; `display` values are worked out from the source
 * and the formatting around it, and are never written into a block's XML (`docx/newLists` reads a
 * list reference off one to decide what numbering.xml needs); `session` values are what export
 * reads but neither writes nor derives again - which block, control or link this is, whether a
 * comment came in with the file, the name of a preserved element. `sourceEquality` compares source
 * and session values so a display refresh does not rebuild an untouched block.
 *
 * Its class records provenance: `preserved` holds XML, `derived` is calculated from other data,
 * `identity` identifies content (imported or allocated here), and `model` holds editable values.
 * Classification alone does not promise a public API; the plugin guide defines that boundary.
 *
 * Lock flags are derived from control XML but have the source role: ignoring them in a comparison
 * would let an unlock pass as a display refresh. The imported flags are derived with the session
 * role because export reads them to decide whether to rewrite comment parts.
 * Both columns are checked against the schema by the adjacent tests.
 */

import { type MarkType, NodeType } from "prosemirror-model";

export type AttrRole = "source" | "display" | "session";

export type AttrClass = "preserved" | "derived" | "identity" | "model";

/** What is declared about one attr: what a comparison does with it, and where its value comes from */
export type AttrFacts = { readonly role: AttrRole; readonly class: AttrClass };

type AttrTable = Readonly<Record<string, Readonly<Record<string, AttrFacts>>>>;

export const NODE_ATTR_ROLES: AttrTable = {
  doc: {
    // The definitions of the lists started while editing. The export writes numbering.xml from
    // them and nothing derives them again, so they are a source value the document node holds
    newLists: { role: "source", class: "model" },
    // The section that closes the body, as the document wrote it. The export writes it back out
    // between the blocks and the tail, so a change to it is a change to the document
    sectPr: { role: "source", class: "preserved" },
    // What each side story currently says. The comment writer puts a comment's back into the
    // Comments part, so a change to one is a change to the document
    stories: { role: "source", class: "model" },
  },
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
    styleConditions: { role: "display", class: "derived" },
    styleBands: { role: "display", class: "derived" },
    // The markers a table and a row carry are content the writer puts back exactly where it
    // stood, so a change to one is a change to the table rather than a display refresh
    leadingXml: { role: "source", class: "preserved" },
  },
  tableRow: {
    trAttrs: { role: "source", class: "preserved" },
    tblPrEx: { role: "source", class: "preserved" },
    trPr: { role: "source", class: "preserved" },
    format: { role: "display", class: "derived" },
    leadingXml: { role: "source", class: "preserved" },
    trailingXml: { role: "source", class: "preserved" },
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
    trailingXml: { role: "source", class: "preserved" },
  },
  rawBlock: {
    xml: { role: "source", class: "preserved" },
    srcId: { role: "session", class: "identity" },
    name: { role: "session", class: "identity" },
    display: { role: "session", class: "derived" },
    guarded: { role: "session", class: "derived" },
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
    paraId: { role: "source", class: "identity" },
    resolved: { role: "source", class: "model" },
    extensionXml: { role: "source", class: "preserved" },
    replies: { role: "source", class: "model" },
    // Read to decide whether the extended comment part is rewritten at all; it is never written
    threadImported: { role: "session", class: "derived" },
  },
  noteReference: {
    kind: { role: "source", class: "model" },
    id: { role: "source", class: "identity" },
    customMarkFollows: { role: "source", class: "model" },
    referenceXml: { role: "source", class: "preserved" },
    // Worked out from the notes part as the document was opened, and the writer puts back the
    // reference alone: a note renumbered around an edit is the same reference it was
    label: { role: "display", class: "derived" },
  },
  rawRunContent: {
    xml: { role: "source", class: "preserved" },
    // What the fragment is and how it is drawn are read off the preservation table when the
    // document is opened and never worked out again, so they travel with the fragment itself
    element: { role: "session", class: "identity" },
    display: { role: "session", class: "derived" },
    text: { role: "session", class: "derived" },
    guarded: { role: "session", class: "derived" },
  },
  rawInline: {
    xml: { role: "source", class: "preserved" },
    element: { role: "session", class: "identity" },
    display: { role: "session", class: "derived" },
    text: { role: "session", class: "derived" },
    guarded: { role: "session", class: "derived" },
  },
};

/**
 * What every wrapper mark carries (`./wrappers`). The export reads both to put the wrappers back
 * where they stood and writes neither into the file, and neither is worked out again from what
 * stands around it.
 */
const WRAPPER_ATTR_ROLES: Readonly<Record<string, AttrFacts>> = {
  // Which wrapper of its kind this is, counted through the document as it was opened
  key: { role: "session", class: "identity" },
  // How deep inside the other wrappers this one stood in the file
  depth: { role: "session", class: "identity" },
};

export const MARK_ATTR_ROLES: AttrTable = {
  run: {
    rPr: { role: "source", class: "preserved" },
    rAttrs: { role: "source", class: "preserved" },
    format: { role: "display", class: "derived" },
  },
  sdt: {
    sdtPrefix: { role: "source", class: "preserved" },
    ...WRAPPER_ATTR_ROLES,
    contentsLocked: { role: "source", class: "derived" },
    deletionLocked: { role: "source", class: "derived" },
  },
  link: {
    linkPrefix: { role: "source", class: "preserved" },
    href: { role: "source", class: "model" },
    ...WRAPPER_ATTR_ROLES,
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
