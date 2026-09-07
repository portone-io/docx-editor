/**
 * Swaps out only the style name, the list slot, the indents, the spacing, and the alignment in a
 * paragraph formatting fragment (`<w:pPr>`).
 *
 * When one paragraph is changed, the formatting we do not touch, such as tab stops and borders,
 * must stay as the original text wrote it. So we never rebuild the fragment; we change it one child
 * at a time, and within a child one attribute at a time.
 * The fragment we operated on is read back with the parser used when opening the document to build
 * the display values, and the style the paragraph points at is layered in the same order as it was
 * then. That way a changed paragraph's display values are the same as they would be after saving
 * and reopening.
 */

import type {
  LineSpacing,
  NumberingRef,
  ParagraphAlign,
  ParagraphFormat,
  RunFormat,
} from "../model/format";
import type { LevelIndent } from "../numbering/parseNumbering";
import {
  attrPairs,
  elementXml,
  withoutAttrs,
  type XmlAttr,
} from "../ooxml/element";
import { wName } from "../ooxml/names";
import { setAttr } from "../ooxml/precedence";
import {
  P_PR_ORDER,
  parseProps,
  parsePropsXml,
  renderProps,
  setPropsChild,
} from "../ooxml/props";
import { childByLocalName, localPart } from "../ooxml/xml";
import {
  layerParagraphFormat,
  layerRunFormat,
  NO_STYLES,
  paragraphStyleFormat,
  readParagraphFormat,
  type StyleTable,
} from "./formatting";

/** What to do with the indents */
export type IndentChange =
  /** Leaves the indents the paragraph wrote down as they are */
  | { kind: "keep" }
  /** Changes them to the indents the list level lays down */
  | { kind: "level"; indent: LevelIndent }
  /** Removes only the hanging indent the marker used to sit in (when leaving a list) */
  | { kind: "clearHanging" };

export interface ListChange {
  /** The list slot this paragraph will belong to. null takes it out of the list */
  numbering: NumberingRef | null;
  indent: IndentChange;
}

/** The paragraph formatting we operated on, and the display values read back out of it */
export interface ParagraphProps {
  pPr: string | null;
  format: ParagraphFormat | null;
  /** The character formatting the style the paragraph now wears lays down (`styleRun` in `schema`) */
  styleRun: RunFormat | null;
}

/** The attributes that record the left indent. They are removed together when a new left indent is written */
const LEFT_IND_ATTRS = ["left", "start", "leftChars", "startChars"];

/** Of those, the ones that record it in twips. A value we write goes into whichever one the document already used */
const LEFT_TWIPS_IND_ATTRS = ["left", "start"];

/** The attributes that record the hanging indent, which is where a list marker sits */
const HANGING_IND_ATTRS = ["hanging", "hangingChars"];

/** The attributes that record the first-line indent. A hanging indent overrules one, so the two are swapped in together */
const FIRST_LINE_IND_ATTRS = [
  ...HANGING_IND_ATTRS,
  "firstLine",
  "firstLineChars",
];

/**
 * Reads the fragment we operated on back along the same path used when opening the document.
 *
 * The values of the style the paragraph wears are laid down underneath (the same order as import),
 * which for a paragraph pointing at no style of its own is the document's default paragraph style.
 */
export function readParagraphProps(
  pPr: string | null,
  styles: StyleTable = NO_STYLES,
  defaultStyleId: string | null = null
): ParagraphFormat | null {
  const direct = pPr === null ? null : readParagraphFormat(parsePropsXml(pPr));
  const style = paragraphStyleFormat(pPr, styles, defaultStyleId);
  return layerParagraphFormat(style?.paragraph ?? {}, direct);
}

/** The character formatting the style the paragraph wears lays down, read back the same way */
function readParagraphStyleRun(
  pPr: string | null,
  styles: StyleTable = NO_STYLES,
  defaultStyleId: string | null = null
): RunFormat | null {
  const style = paragraphStyleFormat(pPr, styles, defaultStyleId);
  return layerRunFormat(style?.run ?? {}, null);
}

function numPrXml(ref: NumberingRef | null): string | null {
  if (!ref) return null;
  return elementXml(
    wName("numPr"),
    [],
    [
      elementXml(wName("ilvl"), [[wName("val"), `${ref.ilvl}`]]),
      elementXml(wName("numId"), [[wName("val"), `${ref.numId}`]]),
    ]
  );
}

function levelIndAttrs(indent: LevelIndent): XmlAttr[] {
  const attrs: XmlAttr[] = [];
  if (indent.startTwips !== null)
    attrs.push([wName("left"), `${indent.startTwips}`]);
  if (indent.hangingTwips !== null) {
    attrs.push([wName("hanging"), `${indent.hangingTwips}`]);
  } else if (indent.firstLineTwips !== null) {
    attrs.push([wName("firstLine"), `${indent.firstLineTwips}`]);
  }
  return attrs;
}

/** The indent written out from its attributes. Null when none are left, which removes the child */
function indXml(attrs: readonly XmlAttr[]): string | null {
  return attrs.length === 0 ? null : elementXml(wName("ind"), attrs);
}

/** The new `<w:ind>` fragment. undefined when the intent is to leave it as it is */
function nextIndXml(
  ind: Element | null,
  change: IndentChange
): string | null | undefined {
  if (change.kind === "keep") return undefined;
  if (change.kind === "clearHanging") {
    return indXml(withoutAttrs(attrPairs(ind), HANGING_IND_ATTRS));
  }
  const dropped = [...LEFT_IND_ATTRS, ...FIRST_LINE_IND_ATTRS];
  return indXml([
    ...levelIndAttrs(change.indent),
    ...withoutAttrs(attrPairs(ind), dropped),
  ]);
}

/**
 * The `w:ind` attributes with the left indent moved to `leftTwips`.
 *
 * The value takes the slot the document already used for it, so an indent moved out and back
 * comes out written the way it went in. Zero is written as no left indent at all, so a paragraph
 * that had none to begin with gets none back.
 * The character-unit spellings drop out, because Word lets them override the value we just wrote.
 */
function leftIndAttrs(ind: Element | null, leftTwips: number): XmlAttr[] {
  const original = attrPairs(ind);
  const slot = original.find(([name]) =>
    LEFT_TWIPS_IND_ATTRS.includes(localPart(name))
  )?.[0];
  const written = leftTwips > 0 ? `${leftTwips}` : null;
  // Every other spelling of the left indent goes, so the one value is recorded in one place
  const kept = withoutAttrs(
    original,
    LEFT_IND_ATTRS.filter(
      (name) => slot === undefined || name !== localPart(slot)
    )
  );
  if (slot !== undefined) return setAttr(kept, "ind", localPart(slot), written);
  return written === null ? kept : [[wName("left"), written], ...kept];
}

const EMPTY_P_PR = { tag: "w:pPr", attrs: null, children: [] };

/** What text one child is to be changed to. A null xml removes that child */
type ChildEdit = readonly [name: string, xml: string | null];

/**
 * The result of swapping out a few of the fragment's children.
 * If the fragment could not be made out the result is null, and in that case the caller leaves
 * that paragraph untouched.
 */
function editParagraphProps(
  pPr: string | null,
  styles: StyleTable,
  defaultStyleId: string | null,
  plan: (element: Element | null) => readonly ChildEdit[]
): ParagraphProps | null {
  const props = pPr === null ? EMPTY_P_PR : parseProps(pPr);
  const element = pPr === null ? null : parsePropsXml(pPr);
  if (!props || (pPr !== null && !element)) return null;

  const rendered = renderProps(
    plan(element).reduce(
      (kept, [name, xml]) => setPropsChild(kept, name, xml, P_PR_ORDER),
      props
    )
  );
  const next = rendered === "" ? null : rendered;
  return {
    pPr: next,
    format: readParagraphProps(next, styles, defaultStyleId),
    styleRun: readParagraphStyleRun(next, styles, defaultStyleId),
  };
}

export function withListNumbering(
  pPr: string | null,
  change: ListChange,
  styles: StyleTable = NO_STYLES,
  defaultStyleId: string | null = null
): ParagraphProps | null {
  return editParagraphProps(pPr, styles, defaultStyleId, (element) => {
    const ind = nextIndXml(
      element ? childByLocalName(element, "ind") : null,
      change.indent
    );
    const numPr: ChildEdit = ["numPr", numPrXml(change.numbering)];
    return ind === undefined ? [numPr] : [numPr, ["ind", ind]];
  });
}

/**
 * The result of moving the paragraph's own left indent to `leftTwips`.
 * Everything else the indent recorded - the hanging or first-line indent, the right indent -
 * stays as it was.
 */
export function withLeftIndent(
  pPr: string | null,
  leftTwips: number,
  styles: StyleTable = NO_STYLES,
  defaultStyleId: string | null = null
): ParagraphProps | null {
  return editParagraphProps(pPr, styles, defaultStyleId, (element) => {
    const ind = element ? childByLocalName(element, "ind") : null;
    return [["ind", indXml(leftIndAttrs(ind, leftTwips))]];
  });
}

/** What one attribute of a formatting child becomes */
type AttrEdit = readonly [name: string, value: string];

/**
 * The `w:spacing` attributes with the named ones changed.
 * A changed attribute keeps the slot it sat in and a new one goes on the end, so the line
 * spacing can be set without disturbing the space above and below the paragraph.
 */
function spacingAttrs(
  spacing: Element | null,
  edits: readonly AttrEdit[]
): XmlAttr[] {
  // Spacing has always updated every spelling of an edited local name. Updating only the first
  // would leave the WML value unchanged when an extension attribute with that name comes first.
  const changed = attrPairs(spacing).map(
    ([name, value]): XmlAttr => [
      name,
      edits.find(([edited]) => edited === localPart(name))?.[1] ?? value,
    ]
  );
  return edits.reduce(
    (attrs, [name, value]) => setAttr(attrs, "spacing", name, value),
    changed
  );
}

/**
 * The result of changing the line spacing. The space above and below the paragraph is left as it was.
 * `auto` states the multiple in 240ths of a line; the other rules pin the height down in twips.
 */
export function withLineSpacing(
  pPr: string | null,
  spacing: LineSpacing,
  styles: StyleTable = NO_STYLES,
  defaultStyleId: string | null = null
): ParagraphProps | null {
  const line =
    spacing.rule === "auto"
      ? Math.round(spacing.lines * 240)
      : Math.round(spacing.pt * 20);
  return editParagraphProps(pPr, styles, defaultStyleId, (element) => {
    const current = element ? childByLocalName(element, "spacing") : null;
    const attrs = spacingAttrs(current, [
      ["line", `${line}`],
      ["lineRule", spacing.rule],
    ]);
    return [["spacing", elementXml(wName("spacing"), attrs)]];
  });
}

/**
 * The result of pointing the paragraph at a named style.
 * A null id takes the pStyle away, which is how the default style is worn.
 *
 * Nothing else in the fragment moves: as in Word, a style is applied by naming it and not by
 * copying its values into the paragraph, so the direct formatting written here keeps beating it.
 */
export function withParagraphStyle(
  pPr: string | null,
  styleId: string | null,
  styles: StyleTable = NO_STYLES,
  defaultStyleId: string | null = null
): ParagraphProps | null {
  const pStyle =
    styleId === null
      ? null
      : elementXml(wName("pStyle"), [[wName("val"), styleId]]);
  return editParagraphProps(pPr, styles, defaultStyleId, () => [
    ["pStyle", pStyle],
  ]);
}

/** OOXML writes justified alignment as `both` */
const JC_BY_ALIGN: Record<ParagraphAlign, string> = {
  left: "left",
  center: "center",
  right: "right",
  justify: "both",
};

/**
 * The result of changing the paragraph alignment.
 * One alignment is always on, so we provide no way to withdraw the setting (the same as Word).
 */
export function withParagraphAlign(
  pPr: string | null,
  align: ParagraphAlign,
  styles: StyleTable = NO_STYLES,
  defaultStyleId: string | null = null
): ParagraphProps | null {
  const jc = elementXml(wName("jc"), [[wName("val"), JC_BY_ALIGN[align]]]);
  return editParagraphProps(pPr, styles, defaultStyleId, () => [["jc", jc]]);
}
