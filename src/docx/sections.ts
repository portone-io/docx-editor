/**
 * What each section of a document lays down, read once out of the one `w:sectPr` that says it.
 *
 * A document is a sequence of sections. Every section but the last is closed by the paragraph that
 * ends it, which carries that section's `w:sectPr` inside its own `w:pPr` (§17.6.17); the last one
 * is closed by the body, whose `w:sectPr` stands after the final block (§17.6.18).
 *
 * A fragment is only ever sliced, never rebuilt, so an untouched section goes back out as the
 * bytes it arrived as.
 */

import {
  type Props,
  parseProps,
  parsePropsXml,
  renderElement,
  renderProps,
  setChild,
} from "../ooxml/props";
import { isOnElement } from "../ooxml/units";
import {
  attributeByLocalName,
  childByLocalName,
  elementChildren,
  R_NS,
} from "../ooxml/xml";
import { type PageGeometry, readPageGeometry } from "./pageGeometry";

export type HeaderFooterVariant = "default" | "first" | "even";

/** The variants a `w:headerReference` or a `w:footerReference` may name. §17.6.12, §17.6.5 */
export const HEADER_FOOTER_VARIANTS: readonly HeaderFooterVariant[] = [
  "default",
  "first",
  "even",
];

/** The part each variant of one story kind is drawn from, as the relationship id naming it */
export type HeaderFooterRefs = Readonly<
  Record<HeaderFooterVariant, string | null>
>;

const NO_REFS: HeaderFooterRefs = { default: null, first: null, even: null };

/** What one `w:sectPr` says, read once and shared by pagination, headers and editing */
export interface SectionProperties {
  /**
   * The `<w:sectPr>...</w:sectPr>` fragment exactly as written.
   *
   * "" for a section read out of a parsed document rather than out of text: the only fragments this
   * package writes back are the slices it took (a paragraph's `w:pPr` child, the document node's
   * own attr), so a section read off a DOM has no text of its own to hand back.
   */
  xml: string;
  geometry: PageGeometry;
  headerRefs: HeaderFooterRefs;
  footerRefs: HeaderFooterRefs;
  titlePg: boolean;
  /** §17.6.19 `w:pgNumType/@w:start`; null where the section carries the numbering on */
  pageNumberStart: number | null;
  /** §17.6.22 `w:type/@w:val`; null when omitted (nextPage) */
  type:
    | "continuous"
    | "evenPage"
    | "nextPage"
    | "oddPage"
    | "nextColumn"
    | null;
}

const SECTION_TYPES: readonly NonNullable<SectionProperties["type"]>[] = [
  "continuous",
  "evenPage",
  "nextPage",
  "oddPage",
  "nextColumn",
];

function sectionType(sectPr: Element): SectionProperties["type"] {
  const declared = childByLocalName(sectPr, "type");
  const value = declared ? attributeByLocalName(declared, "val") : null;
  return SECTION_TYPES.find((known) => known === value) ?? null;
}

/**
 * The page number this section starts at, and null for one carrying the numbering on.
 *
 * What is checked is the number that arrived rather than the shape it was written in, so a start
 * no page counter could reach reads as no start at all.
 */
function pageNumberStart(sectPr: Element): number | null {
  const pgNumType = childByLocalName(sectPr, "pgNumType");
  const declared = pgNumType ? attributeByLocalName(pgNumType, "start") : null;
  if (declared === null || declared.trim() === "") return null;
  const start = Number(declared);
  return Number.isSafeInteger(start) && start >= 0 ? start : null;
}

/**
 * The relationship id each variant of one story kind is drawn from.
 *
 * A section names each variant once, so where a document names one twice the first reference is
 * the one read, and a reference naming no part is no reference at all.
 */
function storyRefs(
  sectPr: Element,
  kind: "header" | "footer"
): HeaderFooterRefs {
  const refs: Record<HeaderFooterVariant, string | null> = { ...NO_REFS };
  for (const reference of elementChildren(sectPr)) {
    if (reference.localName !== `${kind}Reference`) continue;
    const declared = attributeByLocalName(reference, "type") ?? "default";
    const variant = HEADER_FOOTER_VARIANTS.find((known) => known === declared);
    if (variant === undefined || refs[variant] !== null) continue;
    const id =
      reference.getAttributeNS(R_NS, "id") ??
      attributeByLocalName(reference, "id");
    if (id) refs[variant] = id;
  }
  return refs;
}

/**
 * Everything one `w:sectPr` lays down.
 *
 * The reading is deliberately forgiving: it runs against a document already opened, and a value
 * that cannot be read is a reason to fall back on what a section saying nothing lays down, never a
 * reason to refuse the file.
 */
export function readSectionProperties(sectPr: Element): SectionProperties {
  return {
    xml: "",
    geometry: readPageGeometry(sectPr),
    headerRefs: storyRefs(sectPr, "header"),
    footerRefs: storyRefs(sectPr, "footer"),
    titlePg: isOnElement(childByLocalName(sectPr, "titlePg")),
    pageNumberStart: pageNumberStart(sectPr),
    type: sectionType(sectPr),
  };
}

/**
 * The same, off a fragment held as text, which keeps the text as `xml`.
 *
 * null for anything that is not a readable `w:sectPr`, which leaves the caller holding the
 * fragment it already has rather than a section read out of a guess.
 */
export function parseSectionProperties(xml: string): SectionProperties | null {
  const el = parsePropsXml(xml);
  if (el?.localName !== "sectPr") return null;
  return { ...readSectionProperties(el), xml };
}

/**
 * The first `w:sectPr` of an already parsed body: the one closing the first section, wherever it
 * stands, and the body's own where the document holds a single section.
 */
export function firstSectPrElement(body: Element): Element | null {
  for (const el of body.getElementsByTagName("*")) {
    if (el.localName === "sectPr") return el;
  }
  return null;
}

/** The `w:sectPr` this paragraph's properties carry, and null for a paragraph that ends no section */
export function sectionBreakOf(pPr: string | null): string | null {
  if (pPr === null) return null;
  const props = parseProps(pPr);
  return props?.children.find((child) => child.name === "sectPr")?.xml ?? null;
}

/**
 * The paragraph properties without the section break, and null for properties left holding nothing
 * at all. Properties carrying no break, and properties whose shape cannot be made out, are handed
 * back as they stand.
 *
 * `parseProps` lists direct children alone, so the `w:sectPr` inside a `w:pPrChange` - the
 * properties this paragraph wore before a tracked change - is not this paragraph's own break and
 * stays where it stands.
 */
export function withoutSectionBreak(pPr: string | null): string | null {
  if (pPr === null) return null;
  const props = parseProps(pPr);
  if (props === null) return pPr;
  const without = setChild(props, "sectPr", null);
  if (without.children.length === props.children.length) return pPr;
  return renderProps(without) || null;
}

/** The paragraph properties with this section break in the spot `CT_PPr` gives it */
export function withSectionBreak(pPr: string | null, sectPr: string): string {
  const props = (pPr === null ? null : parseProps(pPr)) ?? EMPTY_P_PR;
  return renderElement(setChild(props, "sectPr", sectPr));
}

const EMPTY_P_PR: Props = { tag: "w:pPr", attrs: null, children: [] };
