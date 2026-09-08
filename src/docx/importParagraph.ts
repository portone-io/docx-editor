/**
 * Moves a single `<w:p>` into an editable paragraph node.
 *
 * Body paragraphs and paragraphs inside table cells take the same path.
 *
 * A `w:sdt` content control and a `w:hyperlink` standing inside the paragraph are both unwrapped, so
 * the text they hold stays editable, and the wrapper each came in rides along on a mark to go back
 * out around the same text. A control may hold a link; a link holding a control is a nesting the
 * marks cannot record in that order, and that control stays whole as one preserved fragment.
 *
 * Nothing here demotes a paragraph. What the editor has no model for - a field character, a
 * tracked insertion, a symbol, a drawing nobody could read - is kept where it stood, inside its
 * run or beside it, by the rule `docx/importPolicy` gives it for the level it stands at. The
 * paragraph around it stays editable, which is what a document full of `w:lastRenderedPageBreak`
 * needs.
 */

import type { Mark, Node as PMNode } from "prosemirror-model";
import { readDrawingPicture } from "../ooxml/image";
import { ST_OnOff } from "../ooxml/simpleTypes";
import {
  attributeByLocalName,
  attrString,
  childByLocalName,
  elementChildren,
  serializeXml,
} from "../ooxml/xml";
import { docxSchema } from "../schema";
// The two comment modules are named outright rather than through the folder's barrel: the barrel
// also carries the writer, which reads a story back out (`./story`), and a story is read here
import { commentParaId, importedCommentReplies } from "./comments/model";
import { type ImportedComments, NO_COMMENTS } from "./comments/reading";
import { readParagraphFormat, readRunFormat } from "./formatting";
import {
  type LinkTargets,
  NO_LINK_TARGETS,
  readHyperlinkWrapper,
} from "./hyperlink";
import { policyFor } from "./importPolicy";
import {
  buildPreservedInline,
  buildPreservedRunContent,
  preservationOf,
} from "./importPreserved";
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
 * Moves a `w:drawing` into an image node. null for a drawing we do not interpret, which leaves it
 * to be kept whole inside its run.
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

/**
 * Moves a single element inside a run into an inline node. null for one this reader has no node
 * for, and for one it has a node for but could not read: an unreadable drawing, an annotation
 * reference naming nothing.
 */
function buildModelledRunChild(
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
      const written = attributeByLocalName(el, "customMarkFollows");
      // The attribute is absent far more often than it is there, and only then does it say nothing
      const customMarkFollows =
        written !== null && (ST_OnOff.parse(written) ?? true);
      return [
        docxSchema.nodes.noteReference.create(
          {
            kind,
            id,
            label: note ? noteLabel(kind, id) : "?",
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
): PMNode[] {
  const mark = runMark(run, themeFonts);
  const marks = wrappers.reduce<readonly Mark[]>(
    (set, wrapper) => wrapper.addToSet(set),
    [mark]
  );
  const nodes: PMNode[] = [];
  for (const child of elementChildren(run)) {
    const policy = policyFor(child, "r");
    if (policy.tier === "model") {
      if (child.localName === "rPr") continue;
      const built = buildModelledRunChild(
        child,
        marks,
        images,
        comments,
        notes,
        noteLabel
      );
      if (built !== null) {
        nodes.push(...built);
        continue;
      }
    }
    nodes.push(
      buildPreservedRunContent(child, marks, preservationOf(policy, "r"))
    );
  }
  // A run with no characters at all has nowhere to attach formatting, so we hold on to its original XML as is
  if (nodes.length === 0) {
    return [
      docxSchema.nodes.rawInline.create(
        { xml: serializeXml(run), element: run.localName },
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
 * Moves one child of a wrapper into inline nodes, or null for one this reader cannot unwrap: a
 * control inside another wrapper, a link where one may not stand, a marker naming no comment.
 */
function buildModelledWrapperChild(
  child: Element,
  sources: ImportSources,
  wrappers: readonly Mark[],
  linkAllowed: boolean
): PMNode[] | null {
  if (child.localName === "r") {
    return buildRunNodes(
      child,
      sources.images,
      sources.themeFonts,
      sources.comments,
      sources.notes,
      sources.noteLabel,
      wrappers
    );
  }
  if (child.localName === "hyperlink") {
    return linkAllowed ? buildHyperlinkNodes(child, sources, wrappers) : null;
  }
  const commentMarker = commentRangeNode(child, wrappers);
  return commentMarker === null ? null : [commentMarker];
}

/**
 * Moves what stands inside a wrapper - a content control or a hyperlink - into inline nodes that
 * stay editable, each wearing the marks that remember the wrappers they came out of.
 *
 * Anything this reader cannot unwrap stays whole where it stood, so the wrapper and the text
 * beside it are still editable. null only for a wrapper holding nothing at all, which leaves the
 * marks nothing to hang on.
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
    const policy = policyFor(child, "wrapper");
    if (policy.tier === "model") {
      const built = buildModelledWrapperChild(
        child,
        sources,
        wrappers,
        linkAllowed
      );
      if (built !== null) {
        nodes.push(...built);
        continue;
      }
    }
    nodes.push(
      buildPreservedInline(child, wrappers, preservationOf(policy, "wrapper"))
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
    key: nextKey(linkCounts, el),
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
    key: nextKey(controlCounts, el),
    contentsLocked: wrapper.contentsLocked,
    deletionLocked: wrapper.deletionLocked,
  });
  return buildWrappedNodes(wrapper.content, sources, [mark], true);
}

/**
 * Moves one child of a paragraph into inline nodes, or null for one this reader cannot take
 * apart: a control or a link it could not put back together, a marker naming no comment.
 */
function buildModelledParagraphChild(
  child: Element,
  sources: ImportSources
): PMNode[] | null {
  if (child.localName === "r") {
    return buildRunNodes(
      child,
      sources.images,
      sources.themeFonts,
      sources.comments,
      sources.notes,
      sources.noteLabel
    );
  }
  if (child.localName === "sdt") return buildSdtNodes(child, sources);
  if (child.localName === "hyperlink") {
    return buildHyperlinkNodes(child, sources, []);
  }
  const commentMarker = commentRangeNode(child);
  return commentMarker === null ? null : [commentMarker];
}

/**
 * Moves a paragraph into an editable node.
 *
 * Every paragraph opens editable. A child the reader has no model for, and a control or a link
 * whose shape it could not put back together, is kept whole where it stood instead of standing
 * the whole paragraph down.
 */
export function buildParagraph(
  el: Element,
  srcId: string | null,
  sources: ImportSources = NO_IMPORT_SOURCES
): PMNode {
  const inline: PMNode[] = [];
  let pPrElement: Element | null = null;
  for (const child of elementChildren(el)) {
    const policy = policyFor(child, "p");
    if (policy.tier === "model") {
      if (child.localName === "pPr") {
        pPrElement = child;
        continue;
      }
      const built = buildModelledParagraphChild(child, sources);
      if (built !== null) {
        inline.push(...built);
        continue;
      }
    }
    inline.push(buildPreservedInline(child, [], preservationOf(policy, "p")));
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
