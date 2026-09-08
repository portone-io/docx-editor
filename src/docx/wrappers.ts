/**
 * The inline wrappers this editor takes apart and puts back together, one entry each.
 *
 * A wrapper is an element holding inline content that goes back out around the same content: a
 * content control (`w:sdt`, §17.5.2.17), a hyperlink (`w:hyperlink`, §17.16.22). `EG_PContent`
 * lets them nest in any order and any depth, so import reads the nesting as the file wrote it and
 * export rebuilds it from the marks (`schema/wrappers` holds the order, `docx/importParagraph`
 * reads, `docx/serializeParagraph` writes).
 *
 * A kind knows three things and nothing about either walk: what a mark reads out of one element,
 * what opens it again, and what closes it. Tracked changes (`w:ins`, `w:del`) and simple fields
 * (`w:fldSimple`) are wrappers of the same shape and arrive as one entry here beside one mark spec.
 *
 * The opening tag is the verbatim string the file wrote wherever there is one, so a wrapper nobody
 * edited goes back out byte for byte; only the closing tag is written from scratch.
 */

import type { Mark } from "prosemirror-model";
import { DocxExportError } from "../ooxml/errors";
import { docxSchema } from "../schema";
import type { ExportRefs } from "./exportRefs";
import { readHyperlinkWrapper, relIdIn, withRelId } from "./hyperlink";
import type { ImportSources } from "./importParagraph";
import { copiedControlPrefix, newControlId, readSdtWrapper } from "./sdt";

/** A wrapper read off the file: the mark its content wears, and the element that content stands in */
export interface WrapperReading {
  mark: Mark;
  content: Element;
}

/** One kind of inline wrapper, as import reads it and export writes it */
export interface WrapperKind {
  /** The mark name in the schema */
  readonly mark: string;
  /** The local name of the element it comes from */
  readonly element: string;
  /**
   * The mark this element goes on and the content to read on inside it. null for a shape we do
   * not write back ourselves, which leaves the element preserved whole where it stood.
   */
  read(
    el: Element,
    depth: number,
    sources: ImportSources
  ): WrapperReading | null;
  /** The opening XML this mark goes back out as */
  open(mark: Mark, refs: ExportRefs): string;
  /** What closes it again */
  close(mark: Mark): string;
  /**
   * The same wrapper under a name of its own, for a second stretch claiming one name
   * (`docx/identities`). Absent for a kind that names nothing a document must keep unique.
   */
  copy?(mark: Mark): Mark;
}

/**
 * Hands out the number that tells one wrapper of a kind from the next.
 *
 * Each count runs per parsed document and per kind, so reading a file twice hands out the same
 * numbers and the first control, like the first link, always gets 0. A wrapper only has to be told
 * from the others of its own kind, and a wrapper made during editing needs no number of its own.
 */
const counts = new WeakMap<Document, Map<string, number>>();

function nextKey(kind: WrapperKind, el: Element): number {
  const perDocument = counts.get(el.ownerDocument) ?? new Map<string, number>();
  counts.set(el.ownerDocument, perDocument);
  const key = perDocument.get(kind.mark) ?? 0;
  perDocument.set(kind.mark, key + 1);
  return key;
}

/**
 * The content control. Its content stands in a `w:sdtContent` of its own, which is the tag this
 * writes rather than preserves.
 */
const SDT: WrapperKind = {
  mark: "sdt",
  element: "sdt",
  read(el, depth) {
    const wrapper = readSdtWrapper(el);
    if (!wrapper) return null;
    return {
      mark: docxSchema.marks.sdt.create({
        sdtPrefix: wrapper.prefix,
        depth,
        key: nextKey(SDT, el),
        contentsLocked: wrapper.contentsLocked,
        deletionLocked: wrapper.deletionLocked,
      }),
      content: wrapper.content,
    };
  },
  open(mark) {
    const prefix: unknown = mark.attrs.sdtPrefix;
    if (typeof prefix !== "string") {
      throw new DocxExportError(
        "lost-original",
        "a content control has lost the opening XML it goes back out as"
      );
    }
    return `${prefix}<w:sdtContent>`;
  },
  close: () => "</w:sdtContent></w:sdt>",
  copy(mark) {
    const prefix: unknown = mark.attrs.sdtPrefix;
    if (typeof prefix !== "string") return mark;
    return mark.type.create({
      ...mark.attrs,
      sdtPrefix: copiedControlPrefix(prefix, newControlId()),
    });
  },
};

/**
 * The hyperlink. The address comes off the relationship the wrapper names; a link that names a
 * bookmark alone, or one whose relationship leads nowhere we follow, keeps its wrapper and offers
 * no address.
 *
 * A link that came in with an address goes out pointing at the relationship that address now lives
 * on, which for a link nobody retargeted is the one it arrived on: the tag is then the very string
 * it came as. A link the editor made has no tag of its own and gets the smallest one Word reads.
 * With no relationships to hand out - a serializer running outside an export - an imported link
 * still goes out as it came, and one made here has nothing to point at.
 */
const LINK: WrapperKind = {
  mark: "link",
  element: "hyperlink",
  read(el, depth, sources) {
    const wrapper = readHyperlinkWrapper(el);
    if (!wrapper) return null;
    return {
      mark: docxSchema.marks.link.create({
        linkPrefix: wrapper.prefix,
        href:
          wrapper.relId === null
            ? null
            : (sources.links.get(wrapper.relId) ?? null),
        depth,
        key: nextKey(LINK, el),
      }),
      content: el,
    };
  },
  open(mark, refs) {
    const prefix: unknown = mark.attrs.linkPrefix;
    const original = typeof prefix === "string" ? prefix : null;
    const href: unknown = mark.attrs.href;
    // A link naming a bookmark alone carries no address, and its wrapper says where it goes
    if (typeof href !== "string") {
      if (original !== null) return original;
      throw new DocxExportError(
        "lost-original",
        "a hyperlink carries neither an address nor the opening XML it goes back out as"
      );
    }
    const relId = refs.links.relIdOf(
      href,
      original === null ? null : relIdIn(original)
    );
    if (relId === undefined) {
      if (original !== null) return original;
      throw new DocxExportError(
        "unsupported-content",
        "an inserted hyperlink has no relationship to point at; export it through exportDocx"
      );
    }
    return withRelId(original ?? "<w:hyperlink>", relId);
  },
  close: () => "</w:hyperlink>",
};

/**
 * Every wrapper the editor models, in the order the schema declares their marks: what stands
 * first here stands outside the others where two of them are written at one depth.
 */
export const WRAPPER_KINDS: readonly WrapperKind[] = [SDT, LINK];

/** The kind this element is one of, undefined for an element no kind reads */
export function wrapperKindFor(el: Element): WrapperKind | undefined {
  return WRAPPER_KINDS.find((kind) => kind.element === el.localName);
}

/** The kind this mark stands for. A wrapper mark nobody registered has no way back out */
export function wrapperKindOf(mark: Mark): WrapperKind {
  const kind = WRAPPER_KINDS.find((entry) => entry.mark === mark.type.name);
  if (!kind) {
    throw new DocxExportError(
      "unsupported-content",
      `inline wrapper we cannot serialize: ${mark.type.name}`
    );
  }
  return kind;
}
