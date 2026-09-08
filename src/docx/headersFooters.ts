/** Resolves the header and footer stories one section refers to, for the page preview. */

import type { ParagraphAlign } from "../model/format";
import { ALIGN_BY_JC, isOnElement } from "../ooxml/units";
import {
  attributeByLocalName,
  decodeUtf8,
  elementChildren,
  parseXml,
  R_NS,
  W_NS,
} from "../ooxml/xml";
import { runContentText } from "./importPolicy";
import { relatedPartPath } from "./packageParts";
import { readRelationships, relsPathOf, resolveTarget } from "./relationships";
import {
  HEADER_FOOTER_VARIANTS,
  type HeaderFooterRefs,
  type SectionProperties,
} from "./sections";

export type PageField = "PAGE" | "NUMPAGES";

export type HeaderFooterSegment =
  | { kind: "text"; value: string }
  | { kind: "field"; field: PageField };

/** A display-only projection. The original part remains untouched in the package. */
export interface HeaderFooterContent {
  segments: readonly HeaderFooterSegment[];
  align: ParagraphAlign | null;
}

export interface HeaderFooterVariants {
  default: HeaderFooterContent | null;
  first: HeaderFooterContent | null;
  even: HeaderFooterContent | null;
}

export interface HeadersFooters {
  headers: HeaderFooterVariants;
  footers: HeaderFooterVariants;
  firstPageDifferent: boolean;
  evenAndOdd: boolean;
  pageNumberStart: number;
}

const EMPTY_VARIANTS: HeaderFooterVariants = {
  default: null,
  first: null,
  even: null,
};

export const NO_HEADERS_FOOTERS: HeadersFooters = {
  headers: EMPTY_VARIANTS,
  footers: EMPTY_VARIANTS,
  firstPageDifferent: false,
  evenAndOdd: false,
  pageNumberStart: 1,
};

function pageField(instruction: string): PageField | null {
  const name = instruction.trim().split(/\s+/, 1)[0]?.toUpperCase();
  return name === "PAGE" || name === "NUMPAGES" ? name : null;
}

interface ComplexField {
  instruction: string;
  separated: boolean;
  field: PageField | null;
}

function paragraphSegments(paragraph: Element): HeaderFooterSegment[] {
  const segments: HeaderFooterSegment[] = [];
  const fields: ComplexField[] = [];

  const suppressed = (): boolean =>
    fields.some((field) => !field.separated || field.field !== null);
  const text = (value: string) => {
    if (!value || suppressed()) return;
    const previous = segments.at(-1);
    if (previous?.kind === "text") previous.value += value;
    else segments.push({ kind: "text", value });
  };
  const dynamic = (field: PageField) => {
    if (!suppressed()) segments.push({ kind: "field", field });
  };

  const visit = (el: Element): void => {
    if (
      el.namespaceURI === W_NS &&
      (el.localName === "drawing" ||
        el.localName === "pict" ||
        el.localName === "object" ||
        el.localName === "tbl" ||
        el.localName === "txbxContent")
    ) {
      return;
    }
    if (el.namespaceURI === W_NS && el.localName === "fldSimple") {
      const field = pageField(attributeByLocalName(el, "instr") ?? "");
      if (field) dynamic(field);
      else for (const child of elementChildren(el)) visit(child);
      return;
    }

    if (el.namespaceURI === W_NS && el.localName === "fldChar") {
      const kind = attributeByLocalName(el, "fldCharType");
      if (kind === "begin") {
        fields.push({ instruction: "", separated: false, field: null });
      } else if (kind === "separate") {
        const active = fields.at(-1);
        if (active) {
          active.field = pageField(active.instruction);
          active.separated = true;
          if (active.field) {
            // The field itself replaces the cached result that follows w:separate.
            const outerSuppresses = fields
              .slice(0, -1)
              .some((field) => !field.separated || field.field !== null);
            if (!outerSuppresses) {
              segments.push({ kind: "field", field: active.field });
            }
          }
        }
      } else if (kind === "end") {
        const active = fields.pop();
        if (active && !active.separated) {
          const field = pageField(active.instruction);
          if (field) dynamic(field);
        }
      }
      return;
    }

    if (el.namespaceURI === W_NS && el.localName === "instrText") {
      const active = fields.at(-1);
      if (active && !active.separated)
        active.instruction += el.textContent ?? "";
      return;
    }
    // Everything a run child puts on screen is `./importPolicy`'s answer, so a header reads a
    // carriage return and a no-break hyphen as the body does
    const own = runContentText(el);
    if (own !== null) {
      text(own);
      return;
    }
    for (const child of elementChildren(el)) visit(child);
  };

  for (const child of elementChildren(paragraph)) visit(child);
  return segments;
}

function readContent(bytes: Uint8Array): HeaderFooterContent {
  const root = parseXml(decodeUtf8(bytes).text).documentElement;
  const paragraphs = elementChildren(root).filter(
    (child) => child.namespaceURI === W_NS && child.localName === "p"
  );
  const segments: HeaderFooterSegment[] = [];
  paragraphs.forEach((paragraph, index) => {
    if (index > 0) segments.push({ kind: "text", value: "\n" });
    segments.push(...paragraphSegments(paragraph));
  });
  const pPr = paragraphs[0]
    ? elementChildren(paragraphs[0]).find(
        (child) => child.namespaceURI === W_NS && child.localName === "pPr"
      )
    : undefined;
  const jc = pPr
    ? elementChildren(pPr).find(
        (child) => child.namespaceURI === W_NS && child.localName === "jc"
      )
    : undefined;
  return {
    segments,
    align: jc
      ? (ALIGN_BY_JC[attributeByLocalName(jc, "val") ?? ""] ?? null)
      : null,
  };
}

function settingsEvenAndOdd(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): boolean {
  const path = relatedPartPath(parts, mainPartPath, `${R_NS}/settings`);
  const bytes = path === null ? undefined : parts.get(path);
  if (!bytes) return false;
  const root = parseXml(decodeUtf8(bytes).text).documentElement;
  const setting = root
    .getElementsByTagNameNS(W_NS, "evenAndOddHeaders")
    .item(0);
  return isOnElement(setting);
}

/** The story each variant of the section refers to, as far as the package actually holds one */
function readVariants(
  parts: Map<string, Uint8Array>,
  mainPartPath: string,
  refs: HeaderFooterRefs,
  kind: "header" | "footer"
): HeaderFooterVariants {
  const relationships = new Map(
    readRelationships(parts, relsPathOf(mainPartPath)).map((entry) => [
      entry.id,
      entry,
    ])
  );
  const variants: HeaderFooterVariants = { ...EMPTY_VARIANTS };
  for (const variant of HEADER_FOOTER_VARIANTS) {
    const id = refs[variant];
    if (id === null) continue;
    const relationship = relationships.get(id);
    if (
      !relationship ||
      relationship.external ||
      relationship.type !== `${R_NS}/${kind}`
    ) {
      continue;
    }
    const bytes = parts.get(resolveTarget(mainPartPath, relationship.target));
    if (bytes) variants[variant] = readContent(bytes);
  }
  return variants;
}

/**
 * Reads the display stories one section refers to, and the switches that pick between them.
 *
 * The section itself has already been read (`./sections`), so what is left here is resolving each
 * reference to a part of the package. `w:evenAndOddHeaders` is a document-wide setting rather than
 * a section's own, so it is read from settings.xml instead.
 */
export function readHeadersFooters(
  parts: Map<string, Uint8Array>,
  mainPartPath: string,
  props: SectionProperties
): HeadersFooters {
  return {
    headers: readVariants(parts, mainPartPath, props.headerRefs, "header"),
    footers: readVariants(parts, mainPartPath, props.footerRefs, "footer"),
    firstPageDifferent: props.titlePg,
    evenAndOdd: settingsEvenAndOdd(parts, mainPartPath),
    pageNumberStart: props.pageNumberStart ?? 1,
  };
}

export function displayPageNumber(
  headersFooters: HeadersFooters,
  page: number
): number {
  return headersFooters.pageNumberStart + page - 1;
}

function contentForPage(
  variants: HeaderFooterVariants,
  headersFooters: HeadersFooters,
  page: number
): HeaderFooterContent | null {
  if (page === 1 && headersFooters.firstPageDifferent) return variants.first;
  if (
    headersFooters.evenAndOdd &&
    displayPageNumber(headersFooters, page) % 2 === 0
  ) {
    return variants.even;
  }
  return variants.default;
}

/** Resolves the section variant and evaluates PAGE and NUMPAGES for one visual page. */
export function headerFooterText(
  variants: HeaderFooterVariants,
  headersFooters: HeadersFooters,
  page: number,
  totalPages: number
): string | null {
  const content = contentForPage(variants, headersFooters, page);
  if (!content) return null;
  return content.segments
    .map((segment) => {
      if (segment.kind === "text") return segment.value;
      return segment.field === "PAGE"
        ? `${displayPageNumber(headersFooters, page)}`
        : `${totalPages}`;
    })
    .join("");
}

/** Resolves the direct alignment of the first paragraph in the selected story. */
export function headerFooterAlign(
  variants: HeaderFooterVariants,
  headersFooters: HeadersFooters,
  page: number
): ParagraphAlign | null {
  return contentForPage(variants, headersFooters, page)?.align ?? null;
}
