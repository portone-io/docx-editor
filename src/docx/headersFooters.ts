/**
 * The header and footer stories of a document: where each one is written, which section shows it,
 * what it reads as on a given page, and where an edited one goes back out.
 *
 * A header part is a story of the same schema as the body (`./story`): its blocks are read by the
 * same readers and sliced verbatim by the same scanner, so an edit to one rides a transaction and
 * the part it stands in goes back out as the bytes it arrived as until somebody changes it. What
 * the page preview draws is a projection of that story rather than a reading of its own.
 *
 * Which story a section shows is the section's question (`./sections`): every section names its own
 * variants, so the second section of a document draws its own header and not the first section's.
 */

import type { Node as PMNode } from "prosemirror-model";
import { type ParagraphAlign, toParagraphFormat } from "../model/format";
import { NAMESPACES } from "../ooxml/names";
import {
  ensureRootDeclarations,
  type RootDeclarations,
} from "../ooxml/partSplice";
import { isOnElement } from "../ooxml/units";
import { decodeUtf8, encodeUtf8, parseXml, R_NS, W_NS } from "../ooxml/xml";
import { docxSchema } from "../schema";
import { sameSource } from "../schema/sourceEquality";
import { type StoryKey, storyKey, storyNodeOf } from "../schema/stories";
import { NO_EXPORT_REFS } from "./exportRefs";
import { type FieldSpan, fieldSpans, isFieldCharacter } from "./fields";
import { relatedPartPath } from "./packageParts";
import type { PartPlanner } from "./partPlan";
import { readRelationships, relsPathOf, resolveTarget } from "./relationships";
import {
  type DocumentSection,
  HEADER_FOOTER_VARIANTS,
  type HeaderFooterRefs,
} from "./sections";
import { serializeStory } from "./serializeStory";
import type { SessionStore } from "./session";
import {
  type ImportedStory,
  readStory,
  type StoryDeps,
  storyLeafText,
} from "./story";

/** One story a section may show, as the document currently says it */
export interface HeaderFooterContent {
  story: PMNode;
  /** What the first paragraph of the story is aligned to, which is what the preview is drawn with */
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

/**
 * What picking one section's stories takes beyond the section itself.
 *
 * A section names a relationship, a relationship names a part, and a part is where a story stands,
 * so the two ends are joined once when the document is opened. `w:evenAndOddHeaders` is a
 * document-wide setting rather than a section's own (§17.15.1.29), so it rides along here.
 */
export interface HeaderFooterStories {
  /**
   * The story each header or footer relationship of the main part names. The key says which of the
   * two kinds it is, so a `w:headerReference` naming a footer relationship names no story.
   */
  readonly keyByRelId: ReadonlyMap<string, StoryKey>;
  readonly evenAndOdd: boolean;
}

const PART_ROOT = { header: "hdr", footer: "ftr" } as const;

type HeaderFooterKind = keyof typeof PART_ROOT;

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

function kindOf(relationshipType: string): HeaderFooterKind | null {
  if (relationshipType === `${R_NS}/header`) return "header";
  if (relationshipType === `${R_NS}/footer`) return "footer";
  return null;
}

/**
 * Every header and footer part a section of the document names, read as a story apiece, and the
 * relationship each one answers to.
 *
 * A part is read once however many relationships point at it, and a part whose root is not the
 * element its relationship promises is passed over: the reference names nothing this can draw, and
 * the bytes stay where they are. A part no section names is not read at all, so an orphan a
 * producer left behind - which may hold anything, a document this one never draws included - is
 * neither a story of this document nor a reason to refuse the file, and its bytes go back out as
 * they came.
 */
export function readHeaderFooterStories(
  parts: Map<string, Uint8Array>,
  mainPartPath: string,
  referenced: ReadonlySet<string>,
  depsFor: (partPath: string) => StoryDeps
): { stories: readonly ImportedStory[]; refs: HeaderFooterStories } {
  const keyByRelId = new Map<string, StoryKey>();
  const read = new Map<StoryKey, ImportedStory>();
  for (const relationship of readRelationships(
    parts,
    relsPathOf(mainPartPath)
  )) {
    if (!referenced.has(relationship.id)) continue;
    const kind = relationship.external ? null : kindOf(relationship.type);
    if (kind === null) continue;
    const partPath = resolveTarget(mainPartPath, relationship.target);
    const key = storyKey(kind, partPath);
    if (read.has(key)) {
      keyByRelId.set(relationship.id, key);
      continue;
    }
    const bytes = parts.get(partPath);
    if (!bytes) continue;
    const xml = decodeUtf8(bytes).text;
    const root = parseXml(xml).documentElement;
    if (root.namespaceURI !== W_NS || root.localName !== PART_ROOT[kind]) {
      continue;
    }
    read.set(
      key,
      readStory(
        { kind, id: partPath, partPath },
        { el: root, xml },
        depsFor(partPath)
      )
    );
    keyByRelId.set(relationship.id, key);
  }
  return {
    stories: Array.from(read.values()),
    refs: { keyByRelId, evenAndOdd: settingsEvenAndOdd(parts, mainPartPath) },
  };
}

/**
 * What the first paragraph of a story is aligned to, which is what the whole preview is drawn with.
 *
 * Only that paragraph is read: a story whose first paragraph names no alignment is drawn the way
 * the reader's own writing direction lays it out, not the way some later paragraph is aligned.
 */
function firstAlign(story: PMNode): ParagraphAlign | null {
  for (let at = 0; at < story.childCount; at += 1) {
    const block = story.child(at);
    if (block.type !== docxSchema.nodes.paragraph) continue;
    return toParagraphFormat(block.attrs.format)?.align ?? null;
  }
  return null;
}

function variantContent(
  ids: HeaderFooterRefs,
  kind: HeaderFooterKind,
  refs: HeaderFooterStories,
  storyOf: (key: StoryKey) => PMNode | null
): HeaderFooterVariants {
  const variants: HeaderFooterVariants = { ...EMPTY_VARIANTS };
  for (const variant of HEADER_FOOTER_VARIANTS) {
    const id = ids[variant];
    const key = id === null ? undefined : refs.keyByRelId.get(id);
    if (key === undefined || !key.startsWith(`${kind}:`)) continue;
    const story = storyOf(key);
    if (story !== null) variants[variant] = { story, align: firstAlign(story) };
  }
  return variants;
}

/**
 * The stories this section shows and the switches that pick between them.
 *
 * `storyOf` answers with what the document currently says rather than what the package arrived
 * holding, so a header the editor rewrote is the one the preview draws.
 */
export function variantsFor(
  section: DocumentSection,
  refs: HeaderFooterStories,
  storyOf: (key: StoryKey) => PMNode | null
): HeadersFooters {
  return {
    headers: variantContent(section.props.headerRefs, "header", refs, storyOf),
    footers: variantContent(section.props.footerRefs, "footer", refs, storyOf),
    firstPageDifferent: section.props.titlePg,
    evenAndOdd: refs.evenAndOdd,
    pageNumberStart: section.props.pageNumberStart ?? 1,
  };
}

export function displayPageNumber(
  headersFooters: HeadersFooters,
  page: number
): number {
  return headersFooters.pageNumberStart + page - 1;
}

/** The story this visual page shows, and null where the section declares none for it */
export function headerFooterOn(
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

/** The two fields the preview works out for itself; every other field keeps the result it cached */
type PageFieldName = "PAGE" | "NUMPAGES";

function pageField(instruction: string): PageFieldName | null {
  const name = instruction.split(/\s+/, 1)[0]?.toUpperCase();
  return name === "PAGE" || name === "NUMPAGES" ? name : null;
}

/** What one field of a paragraph takes off the screen, both ends of the range included */
interface Hidden {
  readonly field: FieldSpan;
  readonly from: number;
  readonly to: number;
}

/**
 * What each field of a paragraph takes off the screen.
 *
 * A `PAGE` or a `NUMPAGES` field is worked out again here, so the whole of it goes, the result its
 * producer cached along with it. Every other field keeps that cached result and loses only the
 * instruction it was given, since an instruction is what the field was told rather than what it
 * printed. A simple field holds its result inside itself and so hides nothing at all.
 */
function hiddenBy(field: FieldSpan): Hidden | null {
  if (pageField(field.instr) !== null) {
    return { field, from: field.begin, to: field.end };
  }
  if (field.begin === field.end) return null;
  return { field, from: field.begin, to: field.separate ?? field.end };
}

/**
 * Whether a field standing here is one another field takes off the screen.
 *
 * A field's own range covers where it begins, so only the fields around it are asked: a `PAGE`
 * written inside the instruction of an `{ IF }` is text that field was told rather than a number
 * the page carries, and nothing of it is drawn.
 */
function nested(hidden: readonly Hidden[], field: FieldSpan): boolean {
  return hidden.some(
    (other) =>
      other.field !== field &&
      field.begin >= other.from &&
      field.begin <= other.to
  );
}

function paragraphText(
  paragraph: PMNode,
  page: number,
  totalPages: number
): string {
  const fields = fieldSpans(paragraph, 0);
  const hidden = fields.flatMap((field): Hidden[] => {
    const range = hiddenBy(field);
    return range === null ? [] : [range];
  });
  const numbers = new Map(
    fields.flatMap((field): [number, string][] => {
      const name = pageField(field.instr);
      if (name === null || nested(hidden, field)) return [];
      return [[field.begin, name === "PAGE" ? `${page}` : `${totalPages}`]];
    })
  );
  const pieces: string[] = [];
  paragraph.forEach((child, at) => {
    const number = numbers.get(at);
    if (number !== undefined) pieces.push(number);
    if (isFieldCharacter(child)) return;
    if (hidden.some(({ from, to }) => at >= from && at <= to)) return;
    pieces.push(storyLeafText(child));
  });
  return pieces.join("");
}

/**
 * What the story reads as on one visual page: its paragraphs joined by newlines, with `PAGE` and
 * `NUMPAGES` replaced by the numbers this page actually carries.
 *
 * `page` is the number the page shows rather than its place on the sheet, since that is what a
 * `PAGE` field prints (§17.16.5.45).
 */
export function headerFooterText(
  story: PMNode,
  page: number,
  totalPages: number
): string {
  const lines: string[] = [];
  story.forEach((block) => {
    if (block.type === docxSchema.nodes.paragraph) {
      lines.push(paragraphText(block, page, totalPages));
    }
  });
  return lines.join("\n");
}

/**
 * What a header this editor rewrote has to declare: every block it writes is spelled under `w`, so
 * the part's root binds it rather than each block declaring it again. A part that already binds it
 * - which every one Word writes does - is left exactly as it stands.
 */
const HEADER_MARKUP: RootDeclarations = { namespaces: { w: NAMESPACES.w } };

/** Every header and footer part whose story the document no longer says as the package said it */
function rewrittenParts(
  doc: PMNode,
  session: SessionStore
): ReadonlyMap<string, Uint8Array> | null {
  const parts = new Map<string, Uint8Array>();
  for (const [key, imported] of session.stories) {
    if (imported.kind !== "header" && imported.kind !== "footer") continue;
    const current = storyNodeOf(doc, key);
    if (current === null || sameSource(current, imported.doc)) continue;
    const written = serializeStory(current, imported, imported, {
      ...NO_EXPORT_REFS,
      session,
    });
    const hadBom = decodeUtf8(
      session.parts.get(imported.partPath) ?? new Uint8Array()
    ).hadBom;
    parts.set(
      imported.partPath,
      encodeUtf8(ensureRootDeclarations(written, HEADER_MARKUP), hadBom)
    );
  }
  return parts.size === 0 ? null : parts;
}

/**
 * Writes back the header and footer parts an edit changed, and no others.
 *
 * A story nobody touched is not written at all, so a document opened and exported hands every
 * header part back as the bytes it arrived as, and editing one header leaves the rest untouched.
 */
export const headerFooterPlanner: PartPlanner = {
  name: "headers and footers",
  plan: (doc, session) => rewrittenParts(doc, session),
};
