/**
 * Moves a single `<w:p>` into an editable paragraph node.
 *
 * Body paragraphs and paragraphs inside table cells take the same path.
 *
 * A `w:sdt` content control and a `w:hyperlink` standing inside the paragraph are both unwrapped, so
 * the text they hold stays editable, and the wrapper each came in rides along on a mark to go back
 * out around the same text. Either may hold the other and a control may hold a control: the nesting
 * is read as the file wrote it and recorded on the marks (`docx/wrappers`). A link inside a link is
 * the one arrangement no mark can record, and the inner one stays whole as it always has.
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
import { type LinkTargets, NO_LINK_TARGETS } from "./hyperlink";
import { policyFor } from "./importPolicy";
import {
  buildPreservedInline,
  buildPreservedRunContent,
  preservationOf,
} from "./importPreserved";
import type { ImageSources } from "./media";
import { NO_IMAGES } from "./media";
import { type ImportedNotes, NO_NOTES, type NoteKind, noteById } from "./notes/reading";
import { NO_THEME_FONTS, type ThemeFonts } from "./theme";
import { wrapperFits, wrapperKindFor } from "./wrappers";

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
 * Moves one child of a paragraph or of a wrapper into inline nodes, or null for one this reader
 * cannot take apart: a wrapper whose shape it could not put back together, one holding nothing at
 * all, a marker naming no comment.
 *
 * `depth` is the depth a wrapper met here takes, which is one more than the depth of the wrapper
 * whose content is being read, and `wrappers` are the marks of everything it already stands inside.
 */
function buildModelledInline(
  child: Element,
  sources: ImportSources,
  depth: number,
  wrappers: readonly Mark[]
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
  const kind = wrapperKindFor(child);
  if (kind) {
    // A wrapper its own kind cannot hold stays whole where it stood, wearing the wrappers around it
    if (!wrapperFits(kind, wrappers)) return null;
    const reading = kind.read(child, depth, sources);
    if (!reading) return null;
    return buildContent(reading.content, sources, depth + 1, [
      ...wrappers,
      reading.mark,
    ]);
  }
  const commentMarker = commentRangeNode(child, wrappers);
  return commentMarker === null ? null : [commentMarker];
}

/**
 * Moves what stands inside a wrapper into inline nodes that stay editable, each wearing the marks
 * that remember every wrapper they came out of, outermost first.
 *
 * Anything this reader cannot unwrap stays whole where it stood, so the wrappers and the text
 * beside it are still editable. null only for a wrapper holding nothing at all, which leaves the
 * marks nothing to hang on.
 */
function buildContent(
  content: Element,
  sources: ImportSources,
  depth: number,
  wrappers: readonly Mark[]
): PMNode[] | null {
  const nodes: PMNode[] = [];
  for (const child of elementChildren(content)) {
    const policy = policyFor(child, "wrapper");
    if (policy.tier === "model") {
      const built = buildModelledInline(child, sources, depth, wrappers);
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
 * Moves a paragraph into an editable node.
 *
 * Every paragraph opens editable. A child the reader has no model for, and a wrapper whose shape
 * it could not put back together, is kept whole where it stood instead of standing the whole
 * paragraph down.
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
      const built = buildModelledInline(child, sources, 0, []);
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
