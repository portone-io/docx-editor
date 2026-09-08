/**
 * What the comments in a document come to, worked out once per edit: the threads in document
 * order, the lookup a command asking after one of them uses, and the ranges drawn over the text
 * each was written for.
 *
 * All three are read off the same walk. Which comments a document holds is decided by the document
 * and nothing else, so it is a projection (`./documentProjection`) rather than something worked
 * out again for every command asked, every decoration drawn and every render of the list beside
 * the page.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Plugin } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";
import { storyKey, storyOf, storyText } from "../../docx/story";
import { editorClassNames } from "../../styles/classNames";
import {
  type DocumentComment,
  repliesAttr,
  stringAttr,
} from "../commands/comments/model";
import { documentProjection } from "./documentProjection";

interface MarkerPositions {
  starts: Map<string, number>;
  ends: Map<string, number>;
  references: Array<{ node: PMNode; pos: number }>;
}

function markerPositions(doc: PMNode): MarkerPositions {
  const starts = new Map<string, number>();
  const ends = new Map<string, number>();
  const references: Array<{ node: PMNode; pos: number }> = [];
  doc.descendants((node, pos) => {
    const id = stringAttr(node.attrs.id);
    if (id === null) return true;
    if (node.type.name === "commentStart" && !starts.has(id)) {
      starts.set(id, pos);
    } else if (node.type.name === "commentEnd" && !ends.has(id)) {
      ends.set(id, pos);
    } else if (node.type.name === "commentReference") {
      references.push({ node, pos });
    }
    return true;
  });
  return { starts, ends, references };
}

/** What one comment or reply says, read off the story the document holds it in (`docx/story`) */
function bodyText(doc: PMNode, id: string): string {
  return storyText(storyOf(doc, storyKey("comment", id)));
}

/** The comments in document order, including point comments that have no explicit range. */
function commentsIn(doc: PMNode): readonly DocumentComment[] {
  const markers = markerPositions(doc);
  return markers.references.map(({ node, pos }) => {
    const id = stringAttr(node.attrs.id) ?? "";
    const start = markers.starts.get(id);
    const end = markers.ends.get(id);
    const hasRange = start !== undefined && end !== undefined && start < end;
    const point = start ?? end ?? pos;
    return {
      id,
      author: stringAttr(node.attrs.author),
      authorId: stringAttr(node.attrs.authorId),
      initials: stringAttr(node.attrs.initials),
      date: stringAttr(node.attrs.date),
      text: bodyText(doc, id),
      from: hasRange ? start + 1 : point,
      to: hasRange ? end : point,
      referencePos: pos,
      resolved: node.attrs.resolved === true,
      replies: repliesAttr(node.attrs.replies).map((reply) => ({
        id: reply.id,
        author: reply.author,
        authorId: reply.authorId,
        initials: reply.initials,
        date: reply.date,
        text: bodyText(doc, reply.id),
      })),
    };
  });
}

export interface DocumentComments {
  comments: readonly DocumentComment[];
  /**
   * The comments by id. A document can hold the same id twice - a reference copied and pasted
   * carries the one it was written with - and the first of them wins, which is the one a search
   * through the list in document order arrived at.
   */
  byId: ReadonlyMap<string, DocumentComment>;
  /** The stretch each comment was written for, drawn as it stood when the edit landed */
  decorations: DecorationSet;
}

function deriveComments(doc: PMNode): DocumentComments {
  const comments = commentsIn(doc);
  const byId = new Map<string, DocumentComment>();
  for (const comment of comments) {
    if (!byId.has(comment.id)) byId.set(comment.id, comment);
  }
  return {
    comments,
    byId,
    decorations: DecorationSet.create(
      doc,
      comments
        .filter((comment) => comment.from < comment.to)
        .map((comment) =>
          Decoration.inline(comment.from, comment.to, {
            class: editorClassNames.commentRange,
            "data-comment-id": comment.id,
          })
        )
    ),
  };
}

export const commentProjection = documentProjection<DocumentComments>(
  "docxEditorComments",
  deriveComments,
  {
    // The set the last edit derived. Written as a method so that `this` is the plugin holding it
    decorations(state) {
      return this.getState(state)?.decorations;
    },
  }
);

/**
 * What a state built by `createEditor` registers: the comments it goes on to answer, and the
 * ranges drawn over the text identified by the Word comment markers holding them.
 */
export function commentDecorations(): Plugin {
  return commentProjection.plugin;
}
