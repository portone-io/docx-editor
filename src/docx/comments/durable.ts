/**
 * Writes the ids part and the extensible part, where Word records when a comment was written
 * (`spec/notes/comments.md`). Only entries of comments written or deleted here change.
 */

import type { Node as PMNode } from "prosemirror-model";
import { xmlnsDecl } from "../../ooxml/names";
import {
  ensureRootDeclarations,
  type PartChild,
  partRootProblem,
  type RootDeclarations,
  splicePart,
} from "../../ooxml/partSplice";
import { encodeUtf8, localPart } from "../../ooxml/xml";
import { CONTENT_TYPES_PATH } from "../packageParts";
import { declarePart, type PartPlanContext } from "../partPlan";
import type { SessionStore } from "../session";
import { renderCommentDate, renderCommentId } from "./grammar";
import {
  commentReferencesIn,
  currentCommentBodies,
  originalThreadIds,
} from "./model";
import {
  commentsExtensiblePart,
  commentsIdsPart,
  type DurablePart,
  type DurablePartShape,
} from "./parts";

type DurablePartName = DurablePartShape["name"];

const DURABLE_PARTS: readonly DurablePart[] = [
  commentsIdsPart,
  commentsExtensiblePart,
];

interface DurableChange {
  added: readonly PartChild[];
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
        name: commentsIdsPart.shape.localName,
        xml: renderCommentId(paraId, durableId),
      })),
      dropped: deletedKeys,
    },
    commentsExtensible: {
      added: written.map(({ durableId, dateUtc }) => ({
        name: commentsExtensiblePart.shape.localName,
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
  { shape }: DurablePart,
  change: DurableChange,
  session: SessionStore
): boolean {
  if (shape.xmlIn(session) === null) {
    return change.added.length > 0 && session.parts.has(CONTENT_TYPES_PATH);
  }
  return change.added.length > 0 || change.dropped.size > 0;
}

function markupOf(shape: DurablePartShape): RootDeclarations {
  return { namespaces: { [shape.prefix]: shape.namespace } };
}

function childOrderOf(shape: DurablePartShape): readonly string[] {
  return shape.closingChild === undefined
    ? [shape.localName]
    : [shape.localName, shape.closingChild];
}

function partXml(
  shape: DurablePartShape,
  change: DurableChange,
  session: SessionStore
): string {
  const xml = shape.xmlIn(session);
  if (xml === null) {
    const name = `${shape.prefix}:${shape.rootName}`;
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<${name} ${xmlnsDecl(shape.prefix)}>` +
      change.added.map((entry) => entry.xml).join("") +
      `</${name}>`
    );
  }
  return ensureRootDeclarations(
    splicePart(xml, {
      root: shape.rootName,
      keep: ({ attrs }) => {
        const key = attrs?.find(([name]) => localPart(name) === shape.idAttr);
        return key === undefined || !change.dropped.has(key[1]);
      },
      insert: change.added,
      order: childOrderOf(shape),
    }),
    markupOf(shape)
  );
}

function planPart(
  part: DurablePart,
  change: DurableChange,
  session: SessionStore,
  context: PartPlanContext
): ReadonlyMap<string, Uint8Array> {
  if (!partWritten(part, change, session)) return new Map();
  const { xml, hadBom } = part.shape.imported(session);
  const path = declarePart(part, session, context, xml);
  return new Map([
    [path, encodeUtf8(partXml(part.shape, change, session), hadBom)],
  ]);
}

export function planDurableParts(
  doc: PMNode,
  session: SessionStore,
  context: PartPlanContext
): ReadonlyMap<string, Uint8Array> {
  const changes = durableChanges(doc, session);
  return new Map(
    DURABLE_PARTS.flatMap((part) => [
      ...planPart(part, changes[part.shape.name], session, context),
    ])
  );
}

/** A part that arrived and the planner rewrites, as the export invariants read it */
export interface RewrittenDurablePart {
  name: DurablePartName;
  path: string;
  xml: string;
  declarations: RootDeclarations;
  /** Why it cannot be rewritten around its root, or null when it can */
  rootProblem: string | null;
}

export function rewrittenDurableParts(
  doc: PMNode,
  session: SessionStore
): readonly RewrittenDurablePart[] {
  const changes = durableChanges(doc, session);
  return DURABLE_PARTS.flatMap((part) => {
    const { shape } = part;
    const path = part.pathIn(session);
    const xml = shape.xmlIn(session);
    if (path === null || xml === null) return [];
    if (!partWritten(part, changes[shape.name], session)) return [];
    return [
      {
        name: shape.name,
        path,
        xml,
        declarations: markupOf(shape),
        rootProblem: partRootProblem(xml, shape.rootName),
      },
    ];
  });
}
