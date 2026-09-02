/**
 * Moves a single `<w:p>` into an editable paragraph node.
 *
 * Body paragraphs and paragraphs inside table cells take the same path.
 *
 * A `w:sdt` content control and a `w:hyperlink` standing inside the paragraph are both unwrapped, so
 * the text they hold stays editable, and the wrapper each came in rides along on a mark to go back
 * out around the same text. A control may hold a link; a link holding a control is a nesting the
 * marks cannot record in that order, so the paragraph stays preserved.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import { readDrawingPicture } from "../ooxml/image";
import {
  attrString,
  childByLocalName,
  elementChildren,
  serializeXml,
} from "../ooxml/xml";
import { docxSchema } from "../schema";
import {
  commentParaId,
  type ImportedComments,
  importedCommentReplies,
  NO_COMMENTS,
} from "./comments";
import { readParagraphFormat, readRunFormat } from "./formatting";
import {
  type LinkTargets,
  NO_LINK_TARGETS,
  readHyperlinkWrapper,
} from "./hyperlink";
import type { ImageSources } from "./media";
import { NO_IMAGES } from "./media";
import { type ImportedNotes, NO_NOTES, type NoteKind, noteById } from "./notes";
import { readSdtWrapper } from "./sdt";
import { NO_THEME_FONTS, type ThemeFonts } from "./theme";

function runMark(run: Element, themeFonts: ThemeFonts): Mark {
  const rPr = childByLocalName(run, "rPr");
  return docxSchema.marks.run.create({
    rPr: rPr ? serializeXml(rPr) : null,
    rAttrs: attrString(run),
    format: readRunFormat(rPr, themeFonts),
  });
}

/**
 * Moves a `w:drawing` into an image node. null for a drawing we do not interpret, which
 * puts the whole paragraph back on the preservation path it was on before images were
 * read at all.
 *
 * The original XML rides along, so an image nobody touched goes back out as it came.
 */
function buildImage(
  el: Element,
  marks: readonly Mark[],
  images: ImageSources
): PMNode | null {
  const picture = readDrawingPicture(el);
  if (!picture) return null;
  const src = images.get(picture.relId);
  // A relationship we could not follow leaves nothing to draw
  if (src === undefined) return null;
  return docxSchema.nodes.image.create(
    {
      src,
      extent: picture.extent,
      alt: picture.alt,
      xml: serializeXml(el),
    },
    null,
    marks
  );
}

/** Moves a single element inside a run into an inline node. null for an element we do not model */
function buildRunChild(
  el: Element,
  marks: readonly Mark[],
  images: ImageSources,
  comments: ImportedComments,
  notes: ImportedNotes,
  noteLabel: ImportSources["noteLabel"]
): PMNode[] | null {
  switch (el.localName) {
    case "t": {
      const text = el.textContent ?? "";
      return text ? [docxSchema.text(text, marks)] : [];
    }
    case "br":
      return [
        docxSchema.nodes.hardBreak.create(
          { brAttrs: attrString(el) },
          null,
          marks
        ),
      ];
    case "tab":
      return [
        docxSchema.text(
          "\t",
          docxSchema.marks.tab
            .create({ tabAttrs: attrString(el) })
            .addToSet(marks)
        ),
      ];
    case "drawing": {
      const image = buildImage(el, marks, images);
      return image === null ? null : [image];
    }
    case "commentReference": {
      const id = annotationId(el);
      if (id === null) return null;
      const comment = comments.byId.get(id);
      const paraId = comment?.paraId ?? commentParaId(`comment-${id}`);
      return [
        docxSchema.nodes.commentReference.create(
          {
            id,
            referenceXml: serializeXml(el),
            author: comment?.author ?? null,
            authorId: comment?.authorId ?? null,
            initials: comment?.initials ?? null,
            date: comment?.date ?? null,
            text: comment?.text ?? "",
            commentXml: comment?.xml ?? null,
            imported: true,
            paraId,
            resolved: comment?.resolved ?? false,
            extensionXml: comment?.extensionXml ?? null,
            threadImported: true,
            replies: importedCommentReplies(comments, id),
          },
          null,
          marks
        ),
      ];
    }
    case "footnoteReference":
    case "endnoteReference": {
      const id = annotationId(el);
      if (id === null) return null;
      const kind: NoteKind =
        el.localName === "footnoteReference" ? "footnote" : "endnote";
      const note = noteById(notes, kind, id);
      const customMarkFollows = Array.from(el.attributes).some(
        (entry) =>
          entry.localName === "customMarkFollows" &&
          !["0", "false", "off"].includes(entry.value)
      );
      return [
        docxSchema.nodes.noteReference.create(
          {
            kind,
            id,
            label: note ? noteLabel(kind, id) : "?",
            text: note?.text ?? "",
            customMarkFollows,
            referenceXml: serializeXml(el),
          },
          null,
          marks
        ),
      ];
    }
    default:
      return null;
  }
}

/**
 * Moves one run into inline nodes.
 *
 * A run that stood inside a wrapper - a content control, a hyperlink, or both - is handed that
 * wrapper's marks, which every node it yields wears alongside its own formatting.
 */
function buildRunNodes(
  run: Element,
  images: ImageSources,
  themeFonts: ThemeFonts,
  comments: ImportedComments,
  notes: ImportedNotes,
  noteLabel: ImportSources["noteLabel"],
  wrappers: readonly Mark[] = []
): PMNode[] | null {
  const mark = runMark(run, themeFonts);
  const marks = wrappers.reduce<readonly Mark[]>(
    (set, wrapper) => wrapper.addToSet(set),
    [mark]
  );
  const nodes: PMNode[] = [];
  for (const child of elementChildren(run)) {
    if (child.localName === "rPr") continue;
    const built = buildRunChild(
      child,
      marks,
      images,
      comments,
      notes,
      noteLabel
    );
    if (built === null) return null;
    nodes.push(...built);
  }
  // A run with no characters at all has nowhere to attach formatting, so we hold on to its original XML as is
  if (nodes.length === 0) {
    return [
      docxSchema.nodes.rawInline.create(
        { xml: serializeXml(run) },
        null,
        // The wrappers have to close again around this XML on export, so their marks ride along here too
        wrappers
      ),
    ];
  }
  return nodes;
}

/**
 * Hands out the number that tells one wrapper from the next.
 * Each count runs per parsed document, so reading a file twice hands out the same numbers and the
 * first control, like the first link, always gets 0. Controls and links count apart: a wrapper only
 * has to be told from the others of its own kind.
 */
const controlCounts = new WeakMap<Document, number>();
const linkCounts = new WeakMap<Document, number>();

function nextKey(counts: WeakMap<Document, number>, el: Element): number {
  const key = counts.get(el.ownerDocument) ?? 0;
  counts.set(el.ownerDocument, key + 1);
  return key;
}

/**
 * Word drops these into a control on its own, around text it has looked over, and they put no
 * character on screen. Kept exactly as they came, they leave the text beside them editable.
 */
const NOTHING_ON_SCREEN = [
  "bookmarkStart",
  "bookmarkEnd",
  "proofErr",
  "permStart",
  "permEnd",
];

function annotationId(el: Element): string | null {
  return (
    Array.from(el.attributes).find((attribute) => attribute.localName === "id")
      ?.value ?? null
  );
}

function commentRangeNode(
  el: Element,
  marks: readonly Mark[] = []
): PMNode | null {
  const id = annotationId(el);
  if (id === null) return null;
  const type =
    el.localName === "commentRangeStart"
      ? docxSchema.nodes.commentStart
      : el.localName === "commentRangeEnd"
        ? docxSchema.nodes.commentEnd
        : null;
  return type?.create({ id, xml: serializeXml(el) }, null, marks) ?? null;
}

/**
 * What the body is read with: the images its drawings can point at, the fonts a theme reference
 * resolves to, the addresses its links can point at, and the bodies of its comments and notes.
 * They travel together from `docx/importDocx` down through the tables to here.
 */
export interface ImportSources {
  images: ImageSources;
  themeFonts: ThemeFonts;
  links: LinkTargets;
  comments: ImportedComments;
  notes: ImportedNotes;
  noteLabel: (kind: NoteKind, id: string) => string;
}

export const NO_IMPORT_SOURCES: ImportSources = {
  images: NO_IMAGES,
  themeFonts: NO_THEME_FONTS,
  links: NO_LINK_TARGETS,
  comments: NO_COMMENTS,
  notes: NO_NOTES,
  noteLabel: () => "?",
};

/**
 * Moves what stands inside a wrapper - a content control or a hyperlink - into inline nodes that
 * stay editable, each wearing the marks that remember the wrappers they came out of.
 *
 * null for anything we could not put back together as it came: content that shows text of its own
 * without being a run, a nested control, or a link where one may not stand. That leaves the whole
 * paragraph preserved rather than losing the wrapper quietly.
 */
function buildWrappedNodes(
  content: Element,
  sources: ImportSources,
  wrappers: readonly Mark[],
  // A control may hold a link. The other way round the marks would have to record the link outside
  // the control, which their order does not allow
  linkAllowed: boolean
): PMNode[] | null {
  const nodes: PMNode[] = [];
  for (const child of elementChildren(content)) {
    if (child.localName === "r") {
      const built = buildRunNodes(
        child,
        sources.images,
        sources.themeFonts,
        sources.comments,
        sources.notes,
        sources.noteLabel,
        wrappers
      );
      if (built === null) return null;
      nodes.push(...built);
      continue;
    }
    if (child.localName === "hyperlink" && linkAllowed) {
      const built = buildHyperlinkNodes(child, sources, wrappers);
      if (built === null) return null;
      nodes.push(...built);
      continue;
    }
    const commentMarker = commentRangeNode(child, wrappers);
    if (commentMarker) {
      nodes.push(commentMarker);
      continue;
    }
    if (!NOTHING_ON_SCREEN.includes(child.localName)) return null;
    nodes.push(
      docxSchema.nodes.rawInline.create(
        { xml: serializeXml(child) },
        null,
        wrappers
      )
    );
  }
  // With nothing to hang the marks on, the wrapper would have nothing to come back out around
  if (nodes.length === 0) return null;
  return nodes;
}

/**
 * Moves what stands inside a `w:hyperlink` into inline nodes wearing the mark that remembers it.
 *
 * The address comes off the relationship the wrapper names; a link that names a bookmark alone, or
 * one whose relationship leads nowhere we follow, keeps its wrapper and no address.
 */
function buildHyperlinkNodes(
  el: Element,
  sources: ImportSources,
  wrappers: readonly Mark[]
): PMNode[] | null {
  const wrapper = readHyperlinkWrapper(el);
  if (!wrapper) return null;

  const mark = docxSchema.marks.link.create({
    linkPrefix: wrapper.prefix,
    href:
      wrapper.relId === null
        ? null
        : (sources.links.get(wrapper.relId) ?? null),
    linkKey: nextKey(linkCounts, el),
  });
  return buildWrappedNodes(el, sources, [...wrappers, mark], false);
}

/**
 * Moves what stands inside an inline `w:sdt` into inline nodes that stay editable, each wearing
 * the mark that remembers the control they came out of.
 */
function buildSdtNodes(el: Element, sources: ImportSources): PMNode[] | null {
  const wrapper = readSdtWrapper(el);
  if (!wrapper) return null;

  const mark = docxSchema.marks.sdt.create({
    sdtPrefix: wrapper.prefix,
    sdtKey: nextKey(controlCounts, el),
    contentsLocked: wrapper.contentsLocked,
    deletionLocked: wrapper.deletionLocked,
  });
  return buildWrappedNodes(wrapper.content, sources, [mark], true);
}

/** Moves a paragraph into an editable node. null if it holds content we do not model */
export function buildParagraph(
  el: Element,
  srcId: number | null,
  sources: ImportSources = NO_IMPORT_SOURCES
): PMNode | null {
  const { images, themeFonts } = sources;
  const inline: PMNode[] = [];
  let pPrElement: Element | null = null;
  for (const child of elementChildren(el)) {
    if (child.localName === "pPr") {
      pPrElement = child;
      continue;
    }
    if (child.localName === "r") {
      const runNodes = buildRunNodes(
        child,
        images,
        themeFonts,
        sources.comments,
        sources.notes,
        sources.noteLabel
      );
      if (runNodes === null) return null;
      inline.push(...runNodes);
      continue;
    }
    if (child.localName === "sdt") {
      const sdtNodes = buildSdtNodes(child, sources);
      if (sdtNodes === null) return null;
      inline.push(...sdtNodes);
      continue;
    }
    if (child.localName === "hyperlink") {
      const linkNodes = buildHyperlinkNodes(child, sources, []);
      if (linkNodes === null) return null;
      inline.push(...linkNodes);
      continue;
    }
    const commentMarker = commentRangeNode(child);
    if (commentMarker) {
      inline.push(commentMarker);
      continue;
    }
    // A paragraph child that is not a run, such as a bookmark, keeps its original XML right where it was
    inline.push(
      docxSchema.nodes.rawInline.create({ xml: serializeXml(child) })
    );
  }
  return docxSchema.nodes.paragraph.create(
    {
      srcId,
      pAttrs: attrString(el),
      pPr: pPrElement ? serializeXml(pPrElement) : null,
      format: readParagraphFormat(pPrElement),
    },
    inline
  );
}
