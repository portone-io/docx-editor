/**
 * Reads comment and comment-extension package parts.
 */

import {
  decodeUtf8,
  elementChildren,
  parseXml,
  serializeXml,
  W_NS,
} from "../../ooxml/xml";
import { readRelationships, relsPathOf, resolveTarget } from "../relationships";
import { COMMENTS_EXTENDED_REL_TYPE, COMMENTS_REL_TYPE } from "./constants";

function attribute(el: Element, localName: string): string | null {
  return (
    Array.from(el.attributes).find((entry) => entry.localName === localName)
      ?.value ?? null
  );
}

function inlineCommentText(node: Element): string {
  if (node.localName === "t") return node.textContent ?? "";
  if (node.localName === "tab") return "\t";
  if (node.localName === "br" || node.localName === "cr") return "\n";
  return elementChildren(node).map(inlineCommentText).join("");
}

function commentText(el: Element): string {
  const paragraphs = Array.from(el.getElementsByTagNameNS(W_NS, "p"));
  return paragraphs.map(inlineCommentText).join("\n");
}

export interface ImportedComment {
  id: string;
  author: string | null;
  initials: string | null;
  date: string | null;
  text: string;
  xml: string;
  paraId: string | null;
  parentParaId: string | null;
  resolved: boolean;
  extensionXml: string | null;
}

export interface ImportedComments {
  partPath: string | null;
  xml: string | null;
  hadBom: boolean;
  ordered: readonly ImportedComment[];
  byId: ReadonlyMap<string, ImportedComment>;
  repliesByParentId: ReadonlyMap<string, readonly ImportedComment[]>;
  extendedPartPath: string | null;
  extendedXml: string | null;
  extendedHadBom: boolean;
  extendedOrdered: readonly ImportedCommentExtension[];
}

export const NO_COMMENTS: ImportedComments = {
  partPath: null,
  xml: null,
  hadBom: false,
  ordered: [],
  byId: new Map(),
  repliesByParentId: new Map(),
  extendedPartPath: null,
  extendedXml: null,
  extendedHadBom: false,
  extendedOrdered: [],
};

export interface ImportedCommentExtension {
  paraId: string;
  parentParaId: string | null;
  resolved: boolean;
  xml: string;
}

function lastParagraphId(comment: Element): string | null {
  const paragraphs = Array.from(comment.getElementsByTagNameNS(W_NS, "p"));
  return paragraphs.length === 0
    ? null
    : attribute(paragraphs[paragraphs.length - 1], "paraId");
}

function readCommentExtensions(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): {
  partPath: string | null;
  xml: string | null;
  hadBom: boolean;
  byParaId: ReadonlyMap<string, ImportedCommentExtension>;
  ordered: readonly ImportedCommentExtension[];
} {
  const relationship = readRelationships(parts, relsPathOf(mainPartPath)).find(
    (entry) => entry.type === COMMENTS_EXTENDED_REL_TYPE && !entry.external
  );
  if (!relationship) {
    return {
      partPath: null,
      xml: null,
      hadBom: false,
      byParaId: new Map(),
      ordered: [],
    };
  }
  const partPath = resolveTarget(mainPartPath, relationship.target);
  const bytes = parts.get(partPath);
  if (!bytes) {
    return {
      partPath,
      xml: null,
      hadBom: false,
      byParaId: new Map(),
      ordered: [],
    };
  }
  const { text, hadBom } = decodeUtf8(bytes);
  const root = parseXml(text).documentElement;
  const byParaId = new Map<string, ImportedCommentExtension>();
  const ordered: ImportedCommentExtension[] = [];
  for (const el of elementChildren(root)) {
    if (el.localName !== "commentEx") continue;
    const paraId = attribute(el, "paraId");
    if (paraId === null) continue;
    const extension = {
      paraId,
      parentParaId: attribute(el, "paraIdParent"),
      resolved: ["1", "true", "on"].includes(
        attribute(el, "done")?.toLowerCase() ?? ""
      ),
      xml: serializeXml(el),
    };
    ordered.push(extension);
    if (!byParaId.has(paraId)) byParaId.set(paraId, extension);
  }
  return { partPath, xml: text, hadBom, byParaId, ordered };
}

/** Reads the Comments part related from the main document story. */
export function readComments(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): ImportedComments {
  const relationship = readRelationships(parts, relsPathOf(mainPartPath)).find(
    (entry) => entry.type === COMMENTS_REL_TYPE && !entry.external
  );
  if (!relationship) return NO_COMMENTS;

  const partPath = resolveTarget(mainPartPath, relationship.target);
  const bytes = parts.get(partPath);
  if (!bytes) return { ...NO_COMMENTS, partPath };

  const { text, hadBom } = decodeUtf8(bytes);
  const root = parseXml(text).documentElement;
  const extensions = readCommentExtensions(parts, mainPartPath);
  const base = elementChildren(root)
    .filter((el) => el.localName === "comment")
    .flatMap((el) => {
      const id = attribute(el, "id");
      if (id === null) return [];
      const paraId = lastParagraphId(el);
      const extension = paraId ? extensions.byParaId.get(paraId) : undefined;
      return [
        {
          id,
          author: attribute(el, "author"),
          initials: attribute(el, "initials"),
          date: attribute(el, "date"),
          text: commentText(el),
          xml: serializeXml(el),
          paraId,
          parentParaId: extension?.parentParaId ?? null,
          resolved: extension?.resolved ?? false,
          extensionXml: extension?.xml ?? null,
        },
      ];
    });
  const idByParaId = new Map(
    base.flatMap((comment) =>
      comment.paraId === null ? [] : [[comment.paraId, comment.id] as const]
    )
  );
  const ordered = base;
  const byId = new Map<string, ImportedComment>();
  const repliesByParentId = new Map<string, ImportedComment[]>();
  for (const comment of ordered) {
    if (!byId.has(comment.id)) byId.set(comment.id, comment);
    const parentId =
      comment.parentParaId === null
        ? null
        : (idByParaId.get(comment.parentParaId) ?? null);
    if (parentId !== null) {
      const replies = repliesByParentId.get(parentId) ?? [];
      replies.push(comment);
      repliesByParentId.set(parentId, replies);
    }
  }
  return {
    partPath,
    xml: text,
    hadBom,
    ordered,
    byId,
    repliesByParentId,
    extendedPartPath: extensions.partPath,
    extendedXml: extensions.xml,
    extendedHadBom: extensions.hadBom,
    extendedOrdered: extensions.ordered,
  };
}
