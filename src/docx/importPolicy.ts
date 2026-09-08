/**
 * What every WordprocessingML element the readers may meet is worth keeping as, one sub-table per
 * level of the document.
 *
 * A reader used to know a handful of elements and answer `null` for the rest, and that `null`
 * travelled up through the run and the wrapper until a whole paragraph stood as a placeholder. The
 * table below turns that around: every level has a narrowest preservation of its own - a run child
 * stays inside its run, a paragraph child inside its paragraph, a block stays a block - so an
 * element nobody modelled costs the element and nothing around it.
 *
 * Only `tbl` and `tr` have no node to hold a stranger (`table` holds rows, `tableRow` holds
 * cells), so those two levels demote the table instead, and that is the whole of what demotion is
 * now: this default and a structure a reader could not take apart.
 *
 * A level is a content model of `wml.xsd` rather than an element name, because the same name means
 * different things in different places: `sdt` is a control the paragraph reader unwraps, a cell
 * wrapper under `tr`, a block placeholder under the body, and a row wrapper nothing reads. Keys are
 * Clark names (`{namespace}localName`) so that `m:oMath` and `w:sdt` sit in one map.
 */

import { M_NS, W_NS } from "../ooxml/names";
import type { PreservedDisplay } from "../schema";

export type { PreservedDisplay };

export type PreservationTier =
  /** A reader turns it into an editable node */
  | "model"
  /** A range marker: invisible, and its pair and order must survive every edit */
  | "marker"
  /** A trace the producer regenerates: invisible, and an edit may drop it */
  | "ignorable"
  /** A `CT_R` child: kept verbatim inside its run, wearing the run mark */
  | "runContent"
  /** A paragraph child: kept verbatim inside the paragraph */
  | "inline"
  /** A body or cell child: kept verbatim as a block placeholder */
  | "block";

export interface PreservationRule {
  readonly tier: Exclude<PreservationTier, "model">;
  readonly display: PreservedDisplay;
  /** The character a `text` display puts on screen */
  readonly text?: string;
  /**
   * Whether the deletion guard answers for this element: a range marker whose pair must survive,
   * and a field piece whose begin/separate/end order must survive.
   */
  readonly guarded: boolean;
}

export type ElementPolicy = { readonly tier: "model" } | PreservationRule;

/** What a level with no node to hold a stranger answers: the container around it is demoted */
export type DemotionPolicy = { readonly tier: "demote" };

/** The levels a reader walks. Each is a content model of its own in `wml.xsd` */
export type ContentLevel = "body" | "tbl" | "tr" | "tc" | "p" | "wrapper" | "r";

const MODEL: ElementPolicy = { tier: "model" };
const HIDDEN_MARKER: PreservationRule = {
  tier: "marker",
  display: "hidden",
  guarded: true,
};
const IGNORABLE: PreservationRule = {
  tier: "ignorable",
  display: "hidden",
  guarded: false,
};

function chip(tier: "runContent" | "inline" | "block"): PreservationRule {
  return { tier, display: "chip", guarded: false };
}

const RUN_CHIP = chip("runContent");
const INLINE_CHIP = chip("inline");
const BLOCK_CHIP = chip("block");

/** A field piece: a chip, and one the guard counts, since begin/separate/end stand or fall together */
const FIELD_PIECE: PreservationRule = {
  tier: "runContent",
  display: "chip",
  guarded: true,
};

export function clarkName(el: Element): string {
  return `{${el.namespaceURI ?? ""}}${el.localName}`;
}

function wml(localName: string): string {
  return `{${W_NS}}${localName}`;
}

function rules(
  entries: readonly (readonly [readonly string[], ElementPolicy])[]
): ReadonlyMap<string, ElementPolicy> {
  return new Map(
    entries.flatMap(([names, policy]) =>
      names.map((name): [string, ElementPolicy] => [wml(name), policy])
    )
  );
}

/** `EG_RangeMarkupElements` less the two comment markers, which every level reading them models */
const RANGE_MARKERS = [
  "bookmarkStart",
  "bookmarkEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
  "customXmlInsRangeStart",
  "customXmlInsRangeEnd",
  "customXmlDelRangeStart",
  "customXmlDelRangeEnd",
  "customXmlMoveFromRangeStart",
  "customXmlMoveFromRangeEnd",
  "customXmlMoveToRangeStart",
  "customXmlMoveToRangeEnd",
];

/** The markers of `EG_RunLevelElts` that stand outside `EG_RangeMarkupElements` */
const PERMISSION_MARKERS = ["permStart", "permEnd"];

/** `EG_RunLevelElts` containers, which carry content of their own and so are never a marker */
const REVISION_CONTAINERS = ["ins", "del", "moveFrom", "moveTo"];

const MATH = [`{${M_NS}}oMath`, `{${M_NS}}oMathPara`];

/**
 * `EG_RunInnerContent`, plus the run properties the run mark is read from.
 *
 * `lastRenderedPageBreak` is where the producer's own layout put a page boundary last time it
 * drew the document (Part 1 §17.3.3.13), so it is a cache rather than content: kept, invisible,
 * and no reason to refuse an edit that drops it.
 */
const RUN_LEVEL = rules([
  [["rPr", "t", "br", "tab", "drawing"], MODEL],
  [["commentReference", "footnoteReference", "endnoteReference"], MODEL],
  [["lastRenderedPageBreak"], IGNORABLE],
  [["softHyphen"], { tier: "runContent", display: "hidden", guarded: false }],
  [
    ["noBreakHyphen"],
    { tier: "runContent", display: "text", text: "‑", guarded: false },
  ],
  [["cr"], { tier: "runContent", display: "break", guarded: false }],
  // §17.16.18: a complex field is the begin, the instruction and the end read in order, so a
  // piece taken away leaves the rest saying something the file never said
  [["fldChar", "instrText", "delInstrText"], FIELD_PIECE],
  [
    [
      "sym",
      "ptab",
      "pgNum",
      "pict",
      "object",
      "ruby",
      "contentPart",
      "annotationRef",
      "footnoteRef",
      "endnoteRef",
      "separator",
      "continuationSeparator",
      "delText",
      "dayShort",
      "monthShort",
      "yearShort",
      "dayLong",
      "monthLong",
      "yearLong",
    ],
    RUN_CHIP,
  ],
]);

const PARAGRAPH_CHILDREN: readonly (readonly [
  readonly string[],
  ElementPolicy,
])[] = [
  [RANGE_MARKERS, HIDDEN_MARKER],
  [PERMISSION_MARKERS, HIDDEN_MARKER],
  [["proofErr"], IGNORABLE],
  [REVISION_CONTAINERS, INLINE_CHIP],
  [["fldSimple", "smartTag", "customXml", "dir", "bdo", "subDoc"], INLINE_CHIP],
];

/** `CT_P`: its properties, `EG_PContent`, and the markers and containers that group reaches */
const PARAGRAPH_LEVEL: ReadonlyMap<string, ElementPolicy> = new Map([
  ...rules([
    [["pPr", "r", "hyperlink", "sdt"], MODEL],
    [["commentRangeStart", "commentRangeEnd"], MODEL],
    ...PARAGRAPH_CHILDREN,
  ]),
  ...MATH.map((name): [string, ElementPolicy] => [name, INLINE_CHIP]),
]);

/**
 * The same content model read inside a wrapper the paragraph reader unwrapped.
 *
 * A control may hold a link and the marks record that nesting, so a link is unwrapped here too. A
 * control inside a wrapper is a nesting the marks cannot record in that order, and it stays whole
 * as a chip until the wrapper registry (theme 02) gives it a model.
 */
const WRAPPER_LEVEL: ReadonlyMap<string, ElementPolicy> = new Map([
  ...rules([
    [["r", "hyperlink"], MODEL],
    [["commentRangeStart", "commentRangeEnd"], MODEL],
    [["sdt"], INLINE_CHIP],
    ...PARAGRAPH_CHILDREN,
  ]),
  ...MATH.map((name): [string, ElementPolicy] => [name, INLINE_CHIP]),
]);

const BLOCK_CHILDREN: readonly (readonly [readonly string[], ElementPolicy])[] =
  [
    [["p", "tbl"], MODEL],
    [RANGE_MARKERS, HIDDEN_MARKER],
    [["commentRangeStart", "commentRangeEnd"], HIDDEN_MARKER],
    [PERMISSION_MARKERS, HIDDEN_MARKER],
    [["proofErr"], IGNORABLE],
    [REVISION_CONTAINERS, BLOCK_CHIP],
    [["sdt", "customXml", "altChunk"], BLOCK_CHIP],
  ];

/** `CT_Body`: `EG_BlockLevelElts` and the section the body closes with */
const BODY_LEVEL: ReadonlyMap<string, ElementPolicy> = new Map([
  ...rules([...BLOCK_CHILDREN, [["sectPr"], BLOCK_CHIP]]),
  ...MATH.map((name): [string, ElementPolicy] => [name, BLOCK_CHIP]),
]);

/** `CT_Tc`: its properties and the same block content the body takes */
const CELL_LEVEL: ReadonlyMap<string, ElementPolicy> = new Map([
  ...rules([...BLOCK_CHILDREN, [["tcPr"], MODEL]]),
  ...MATH.map((name): [string, ElementPolicy] => [name, BLOCK_CHIP]),
]);

/**
 * `CT_Tbl`: the markers it opens with, its properties, its grid and its rows.
 *
 * A row wrapper and a revision container have no node here, so they are left out and the level's
 * default demotes the table around them.
 */
const TABLE_LEVEL = rules([
  [["tblPr", "tblGrid", "tr"], MODEL],
  [RANGE_MARKERS, HIDDEN_MARKER],
  [["commentRangeStart", "commentRangeEnd"], HIDDEN_MARKER],
  [PERMISSION_MARKERS, HIDDEN_MARKER],
  [["proofErr"], IGNORABLE],
]);

/** `CT_Row`: its properties, its cells, and a control wrapping one cell */
const ROW_LEVEL = rules([
  [["tblPrEx", "trPr", "tc", "sdt"], MODEL],
  [RANGE_MARKERS, HIDDEN_MARKER],
  [["commentRangeStart", "commentRangeEnd"], HIDDEN_MARKER],
  [PERMISSION_MARKERS, HIDDEN_MARKER],
  [["proofErr"], IGNORABLE],
]);

export const WML_POLICY: Readonly<
  Record<ContentLevel, ReadonlyMap<string, ElementPolicy>>
> = {
  r: RUN_LEVEL,
  p: PARAGRAPH_LEVEL,
  wrapper: WRAPPER_LEVEL,
  body: BODY_LEVEL,
  tc: CELL_LEVEL,
  tbl: TABLE_LEVEL,
  tr: ROW_LEVEL,
};

/** The levels that have a node of their own to keep a stranger in */
export type PreservingLevel = Exclude<ContentLevel, "tbl" | "tr">;

/**
 * The narrowest preservation each of those levels has.
 *
 * It is what an element no sub-table names falls to, and what a reader keeps an element under
 * when the table said `model` and it turned out it could not be read after all.
 */
const LEVEL_PRESERVATION: Readonly<Record<PreservingLevel, PreservationRule>> =
  {
    r: RUN_CHIP,
    p: INLINE_CHIP,
    wrapper: INLINE_CHIP,
    body: BLOCK_CHIP,
    tc: BLOCK_CHIP,
  };

export function defaultPreservation(level: PreservingLevel): PreservationRule {
  return LEVEL_PRESERVATION[level];
}

const LEVEL_DEFAULTS: Readonly<
  Record<ContentLevel, ElementPolicy | DemotionPolicy>
> = {
  ...LEVEL_PRESERVATION,
  tbl: { tier: "demote" },
  tr: { tier: "demote" },
};

export function policyFor(
  el: Element,
  level: ContentLevel
): ElementPolicy | DemotionPolicy {
  return WML_POLICY[level].get(clarkName(el)) ?? LEVEL_DEFAULTS[level];
}

/**
 * The text one run child puts on screen, or null for a child that puts nothing there.
 *
 * Everything that reads a story as plain text - a footnote body, a comment body, a header -
 * asks this rather than keeping a vocabulary of its own, so the three of them cannot disagree
 * about what a `w:cr` or a `w:noBreakHyphen` reads as.
 */
export function runContentText(el: Element): string | null {
  const policy = policyFor(el, "r");
  if (policy.tier === "model") {
    switch (el.localName) {
      case "t":
        return el.textContent ?? "";
      case "tab":
        return "\t";
      case "br":
        return "\n";
      // A drawing and the three annotation references draw something that is not text
      default:
        return null;
    }
  }
  if (policy.tier === "demote") return null;
  if (policy.display === "break") return "\n";
  return policy.display === "text" ? (policy.text ?? null) : null;
}
