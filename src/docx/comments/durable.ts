/**
 * Writes the ids part and the extensible part, where Word records when a comment was written
 * (`spec/notes/comments.md`). Only entries of comments written or deleted here change.
 */

import type { Node as PMNode } from "prosemirror-model";
import { NAMESPACES, xmlnsDecl } from "../../ooxml/names";
import {
  ensureRootDeclarations,
  partRootProblem,
  type RootDeclarations,
  splicePart,
} from "../../ooxml/partSplice";
import { parseAttrs, readTag } from "../../ooxml/tagScan";
import { encodeUtf8, localPart } from "../../ooxml/xml";
import { CONTENT_TYPES_PATH } from "../packageParts";
import type { PartPlanContext } from "../partPlan";
import type { StoryPartKind } from "../protectionPolicy";
import { directoryOf } from "../relationships";
import type { SessionStore } from "../session";
import { renderCommentDate, renderCommentId } from "./grammar";
import {
  commentReferencesIn,
  currentCommentBodies,
  originalThreadIds,
} from "./model";
import { commentIdsPart, commentsExtensiblePart } from "./parts";
import type { ImportedCommentPart } from "./reading";

type DurablePartName = "commentsIds" | "commentsExtensible";

interface DurablePartShape {
  /** As an export refusal names it (`ExportPartName`) */
  name: DurablePartName;
  part: StoryPartKind;
  root: string;
  prefix: "w16cid" | "w16cex";
  /** The children the part's type takes, in order */
  order: readonly string[];
  /** The attribute a deletion finds the entries to drop by */
  key: string;
  markup: RootDeclarations;
  imported(session: SessionStore): ImportedCommentPart<unknown>;
}

const IDS_SHAPE: DurablePartShape = {
  name: "commentsIds",
  part: commentIdsPart,
  root: "commentsIds",
  prefix: "w16cid",
  order: ["commentId"],
  key: "paraId",
  markup: { namespaces: { w16cid: NAMESPACES.w16cid } },
  imported: (session) => session.comments.ids,
};

const EXTENSIBLE_SHAPE: DurablePartShape = {
  name: "commentsExtensible",
  part: commentsExtensiblePart,
  root: "commentsExtensible",
  prefix: "w16cex",
  // `CT_CommentsExtensible` closes on an optional extension list ([MS-DOCX] §2.10.3.2)
  order: ["commentExtensible", "extLst"],
  key: "durableId",
  markup: { namespaces: { w16cex: NAMESPACES.w16cex } },
  imported: (session) => session.comments.extensible,
};

const DURABLE_SHAPES: readonly DurablePartShape[] = [
  IDS_SHAPE,
  EXTENSIBLE_SHAPE,
];

interface DurableChange {
  added: readonly { name: string; xml: string }[];
  /** The keys of the entries it gives up */
  dropped: ReadonlySet<string>;
}

type DurableChanges = Readonly<Record<DurablePartName, DurableChange>>;

/**
 * A durable id the ids part did not arrive with belongs to a comment written here. A comment the
 * story stood behind when opened and no longer holds was deleted here; an orphan entry stays.
 */
function durableChanges(doc: PMNode, session: SessionStore): DurableChanges {
  const { comments } = session;
  const current = currentCommentBodies(commentReferencesIn(doc));
  const arrived = new Set(comments.ids.ordered.map((entry) => entry.durableId));
  const written = Array.from(current.values()).flatMap((comment) =>
    comment.durableId === null ||
    comment.dateUtc === null ||
    arrived.has(comment.durableId)
      ? []
      : [
          {
            paraId: comment.paraId,
            durableId: comment.durableId,
            dateUtc: comment.dateUtc,
          },
        ]
  );
  const deleted = Array.from(
    originalThreadIds(comments, session.commentReferenceIds)
  ).filter((id) => !current.has(id));
  const deletedKeys = new Set(
    deleted.flatMap((id) => {
      const paraId = comments.byId.get(id)?.paraId;
      return paraId == null ? [] : [paraId];
    })
  );
  const deletedDurableIds = new Set(
    comments.ids.ordered.flatMap((entry) =>
      deletedKeys.has(entry.paraId) ? [entry.durableId] : []
    )
  );
  return {
    commentsIds: {
      added: written.map(({ paraId, durableId }) => ({
        name: "commentId",
        xml: renderCommentId(paraId, durableId),
      })),
      dropped: deletedKeys,
    },
    commentsExtensible: {
      added: written.map(({ durableId, dateUtc }) => ({
        name: "commentExtensible",
        xml: renderCommentDate(durableId, dateUtc),
      })),
      dropped: deletedDurableIds,
    },
  };
}

/**
 * A new part is left out of a package with no `[Content_Types].xml` to declare it in, rather than
 * refused: the comment then loses only its UTC date and still reads as its `w:date`.
 */
function partWritten(
  shape: DurablePartShape,
  change: DurableChange,
  session: SessionStore
): boolean {
  const { xml } = shape.imported(session);
  if (xml === null) {
    return change.added.length > 0 && session.parts.has(CONTENT_TYPES_PATH);
  }
  return change.added.length > 0 || change.dropped.size > 0;
}

function attributeOf(child: string, name: string): string | null {
  const tag = readTag(child, 0);
  if (tag === null || tag.kind === "close" || tag.kind === "other") return null;
  const attrs = parseAttrs(
    child.slice(tag.nameEnd, tag.end - (tag.kind === "empty" ? 2 : 1))
  );
  return attrs?.find(([attr]) => localPart(attr) === name)?.[1] ?? null;
}

function partXml(
  shape: DurablePartShape,
  change: DurableChange,
  session: SessionStore
): string {
  const { xml } = shape.imported(session);
  if (xml === null) {
    const name = `${shape.prefix}:${shape.root}`;
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<${name} ${xmlnsDecl(shape.prefix)}>` +
      change.added.map((entry) => entry.xml).join("") +
      `</${name}>`
    );
  }
  return ensureRootDeclarations(
    splicePart(xml, {
      root: shape.root,
      keep: (child) => {
        const key = attributeOf(child, shape.key);
        return key === null || !change.dropped.has(key);
      },
      insert: change.added,
      order: shape.order,
    }),
    shape.markup
  );
}

function planPart(
  shape: DurablePartShape,
  change: DurableChange,
  session: SessionStore,
  context: PartPlanContext
): ReadonlyMap<string, Uint8Array> {
  if (!partWritten(shape, change, session)) return new Map();
  const { part } = shape;
  const imported = shape.imported(session);
  const adding = part.pathIn(session) === null;
  const path = part.writePathIn(session);
  if (adding) {
    context.relationships.add({
      type: part.relType,
      target: path.slice(directoryOf(session.mainPartPath).length),
    });
  }
  if (adding || imported.xml === null) {
    context.contentTypes.addOverride(path, part.contentType);
  }
  return new Map([
    [path, encodeUtf8(partXml(shape, change, session), imported.hadBom)],
  ]);
}

export function planDurableParts(
  doc: PMNode,
  session: SessionStore,
  context: PartPlanContext
): ReadonlyMap<string, Uint8Array> {
  const changes = durableChanges(doc, session);
  return new Map(
    DURABLE_SHAPES.flatMap((shape) => [
      ...planPart(shape, changes[shape.name], session, context),
    ])
  );
}

/** A part the planner rewrites, as the export invariants read it */
export interface RewrittenDurablePart {
  name: DurablePartName;
  path: string | null;
  xml: string | null;
  declarations: RootDeclarations;
  /** Why it cannot be rewritten around its root, or null when it can */
  rootProblem: string | null;
}

export function rewrittenDurableParts(
  doc: PMNode,
  session: SessionStore
): readonly RewrittenDurablePart[] {
  const changes = durableChanges(doc, session);
  return DURABLE_SHAPES.flatMap((shape) => {
    if (!partWritten(shape, changes[shape.name], session)) return [];
    const { xml } = shape.imported(session);
    return [
      {
        name: shape.name,
        path: shape.part.pathIn(session),
        xml,
        declarations: shape.markup,
        rootProblem: xml === null ? null : partRootProblem(xml, shape.root),
      },
    ];
  });
}
