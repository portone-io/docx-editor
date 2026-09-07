/**
 * What a node an edit makes out of another node carries over from it.
 *
 * A paragraph built from another one - either half of the two Enter leaves behind, or a copy - used
 * to arrive holding every attr of the original. Two paragraphs then claimed the same `w14:paraId`,
 * and a paragraph-level `w:sectPr` went out twice, which
 * gives the document a section break it never had.
 *
 * `CLONE_POLICIES` is the one place that answers what becomes of each attr, and it answers per
 * side: the half a split leaves standing is not the half it makes, and a free-standing copy is
 * neither. What the answer turns on is whether the attr names the node - an identifier, or the
 * block it was opened from - or describes it. A name may be held by one node only; a description
 * is free to be copied.
 */

import type { Attrs, Node as PMNode } from "prosemirror-model";
import { attrsText, withoutAttrs } from "../ooxml/element";
import { parseProps, renderProps, setChild } from "../ooxml/props";
import { parseAttrs } from "../ooxml/tagScan";

/** Which half of a split the clone becomes, or a free-standing copy */
export type CloneSide = "before" | "after" | "copy";

export interface ClonePolicy {
  /** The attrs a node made from `source` carries when it stands on `side` */
  attrs(source: PMNode, side: CloneSide): Attrs;
}

/**
 * What the clone carries in one attr, out of what the source carried there.
 *
 * `undefined` leaves the attr out of the clone altogether, so the schema's own default stands
 * where the source's value would have. That is how an attr is dropped rather than nulled: `colspan`
 * defaults to 1, and writing null into it would make a cell the grid cannot measure.
 */
type AttrRule = (value: unknown, side: CloneSide) => unknown;

/** The value the source carried, as it stands */
const carried: AttrRule = (value) => value;

/** Nothing, so the clone starts this attr from the schema's default */
const defaulted: AttrRule = () => undefined;

/**
 * The names Word writes a paragraph's own identity under, which belong to the `w14` vocabulary
 * rather than to WordprocessingML itself.
 */
const PARAGRAPH_ID_ATTRS = ["paraId", "textId"] as const;

/**
 * The opening tag's attributes without the paragraph identifiers.
 *
 * A paragraph carrying none goes back out spelled exactly as it arrived, and so does one whose
 * attribute text cannot be made out, which is left alone rather than written again from a guess.
 */
export function withoutParagraphIds(value: unknown): unknown {
  if (typeof value !== "string") return value;
  const attrs = parseAttrs(value);
  if (attrs === null) return value;
  const kept = withoutAttrs(attrs, PARAGRAPH_ID_ATTRS, "w14");
  if (kept.length === attrs.length) return value;
  return kept.length === 0 ? null : attrsText(kept);
}

/**
 * The paragraph properties without the section break, and null for properties left holding nothing
 * at all.
 *
 * `parseProps` lists direct children alone, so the `w:sectPr` inside a `w:pPrChange` - the
 * properties this paragraph wore before a tracked change - is not this paragraph's own break and
 * stays where it stands.
 */
function withoutSectionBreak(value: unknown): unknown {
  if (typeof value !== "string") return value;
  const props = parseProps(value);
  if (props === null) return value;
  const without = setChild(props, "sectPr", null);
  if (without.children.length === props.children.length) return value;
  return renderProps(without) || null;
}

/**
 * Everything a new cell inherits.
 *
 * The content control around a cell (`sdtPrefix`, and the two locks that come out of it) is
 * deliberately left out: a new cell must not quietly come into the document carrying a copy of
 * somebody else's control, let alone that control's lock.
 */
const INHERITED_CELL_ATTRS = ["tcAttrs", "tcPr", "tcW", "format"] as const;

/**
 * Everything a new row inherits. The row height lives inside `trPr`.
 *
 * The property exceptions (`tblPrEx`) stay behind: we never read them, so a new row simply
 * follows the table's own values.
 */
const INHERITED_ROW_ATTRS = ["trAttrs", "trPr", "format"] as const;

/** A rule for each of these attrs, carrying what the source held there */
function inherits(names: readonly string[]): Record<string, AttrRule> {
  return Object.fromEntries(names.map((name) => [name, carried]));
}

/**
 * A policy that puts each attr the source holds through the rule named for it, and through `rest`
 * where the table names none.
 */
function policy(
  rules: Readonly<Record<string, AttrRule>>,
  rest: AttrRule
): ClonePolicy {
  return {
    attrs(source, side) {
      const attrs: Record<string, unknown> = {};
      for (const [name, value] of Object.entries(source.attrs)) {
        const held = (rules[name] ?? rest)(value, side);
        if (held !== undefined) attrs[name] = held;
      }
      return attrs;
    },
  };
}

/** One policy per node type that an edit can make from another node */
export const CLONE_POLICIES: Readonly<
  Record<"paragraph" | "tableCell" | "tableRow", ClonePolicy>
> = {
  paragraph: policy(
    {
      // The half that carries on where the original stood keeps the original's name. The one the
      // edit made is a new paragraph and goes out with none, so no two ever claim the same
      // paragraph identity.
      pAttrs: (value, side) =>
        side === "before" ? value : withoutParagraphIds(value),
      // A paragraph-level `w:sectPr` ends the section it stands in, so it belongs to whichever
      // paragraph ends up last in that section: the second half of a split, and no half of a copy,
      // which is lifted out of the section it was taken from.
      pPr: (value, side) =>
        side === "after" ? value : withoutSectionBreak(value),
      // What names the block this paragraph was opened from. Only the half still standing where
      // that block stood may go on claiming it; `docx/exportDocx` writes the original bytes back
      // for whatever claims one.
      srcId: (value, side) => (side === "before" ? value : undefined),
    },
    // `format` and `styleRun` among them: display values `editor/plugins/styledParagraphs` derives
    // again from the formatting context, which nothing compares and nothing writes out.
    carried
  ),
  tableCell: policy(inherits(INHERITED_CELL_ATTRS), defaulted),
  tableRow: policy(inherits(INHERITED_ROW_ATTRS), defaulted),
};

/**
 * The two halves Enter leaves behind. `before` keeps the text, `after` keeps the paragraph mark's
 * section break
 */
export function splitParagraphAttrs(parent: PMNode): {
  before: Attrs;
  after: Attrs;
} {
  return {
    before: CLONE_POLICIES.paragraph.attrs(parent, "before"),
    after: CLONE_POLICIES.paragraph.attrs(parent, "after"),
  };
}
