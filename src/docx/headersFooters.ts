/** Reads the first section's header and footer stories for the page preview. */

import type { ParagraphAlign } from "../model/format";
import { ALIGN_BY_JC } from "../ooxml/units";
import {
  decodeUtf8,
  elementChildren,
  parseXml,
  R_NS,
  W_NS,
} from "../ooxml/xml";
import { readRelationships, relsPathOf, resolveTarget } from "./relationships";

export type HeaderFooterVariant = "default" | "first" | "even";
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

function attribute(el: Element, localName: string): string | null {
  return (
    Array.from(el.attributes).find((entry) => entry.localName === localName)
      ?.value ?? null
  );
}

function isOn(el: Element | null): boolean {
  if (!el) return false;
  const value = attribute(el, "val")?.toLowerCase();
  return value !== "0" && value !== "false" && value !== "off";
}

function firstSectPrIn(root: Element): Element | null {
  const namespaced = root.getElementsByTagNameNS(W_NS, "sectPr").item(0);
  if (namespaced) return namespaced;
  for (const el of root.getElementsByTagName("*")) {
    if (el.localName === "sectPr") return el;
  }
  return null;
}

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
      const field = pageField(attribute(el, "instr") ?? "");
      if (field) dynamic(field);
      else for (const child of elementChildren(el)) visit(child);
      return;
    }

    if (el.namespaceURI === W_NS && el.localName === "fldChar") {
      const kind = attribute(el, "fldCharType");
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
    if (el.namespaceURI === W_NS && el.localName === "t") {
      text(el.textContent ?? "");
      return;
    }
    if (el.namespaceURI === W_NS && el.localName === "tab") {
      text("\t");
      return;
    }
    if (
      el.namespaceURI === W_NS &&
      (el.localName === "br" || el.localName === "cr")
    ) {
      text("\n");
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
    align: jc ? (ALIGN_BY_JC[attribute(jc, "val") ?? ""] ?? null) : null,
  };
}

function settingsEvenAndOdd(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): boolean {
  const relationship = readRelationships(parts, relsPathOf(mainPartPath)).find(
    (entry) => entry.type === `${R_NS}/settings` && !entry.external
  );
  if (!relationship) return false;
  const bytes = parts.get(resolveTarget(mainPartPath, relationship.target));
  if (!bytes) return false;
  const root = parseXml(decodeUtf8(bytes).text).documentElement;
  const setting = root
    .getElementsByTagNameNS(W_NS, "evenAndOddHeaders")
    .item(0);
  return isOn(setting);
}

function readVariants(
  parts: Map<string, Uint8Array>,
  mainPartPath: string,
  sectPr: Element,
  kind: "header" | "footer"
): HeaderFooterVariants {
  const relationships = new Map(
    readRelationships(parts, relsPathOf(mainPartPath)).map((entry) => [
      entry.id,
      entry,
    ])
  );
  const variants: HeaderFooterVariants = { ...EMPTY_VARIANTS };
  for (const reference of elementChildren(sectPr)) {
    if (reference.localName !== `${kind}Reference`) continue;
    const type = attribute(reference, "type") ?? "default";
    if (type !== "default" && type !== "first" && type !== "even") continue;
    const id =
      reference.getAttributeNS(R_NS, "id") ?? attribute(reference, "id");
    if (!id) continue;
    const relationship = relationships.get(id);
    if (
      !relationship ||
      relationship.external ||
      relationship.type !== `${R_NS}/${kind}`
    ) {
      continue;
    }
    const bytes = parts.get(resolveTarget(mainPartPath, relationship.target));
    if (bytes) variants[type] = readContent(bytes);
  }
  return variants;
}

function pageNumberStart(sectPr: Element): number {
  const pgNumType = elementChildren(sectPr).find(
    (child) => child.localName === "pgNumType"
  );
  const declared = pgNumType ? attribute(pgNumType, "start") : null;
  if (declared === null) return 1;
  const start = Number(declared);
  return Number.isSafeInteger(start) && start >= 0 ? start : 1;
}

/** Reads the first section's display stories and section-level selection switches. */
export function readHeadersFooters(
  parts: Map<string, Uint8Array>,
  mainPartPath: string,
  body: Element
): HeadersFooters {
  const sectPr = firstSectPrIn(body);
  if (!sectPr) return NO_HEADERS_FOOTERS;
  return {
    headers: readVariants(parts, mainPartPath, sectPr, "header"),
    footers: readVariants(parts, mainPartPath, sectPr, "footer"),
    firstPageDifferent: isOn(
      elementChildren(sectPr).find((child) => child.localName === "titlePg") ??
        null
    ),
    evenAndOdd: settingsEvenAndOdd(parts, mainPartPath),
    pageNumberStart: pageNumberStart(sectPr),
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
