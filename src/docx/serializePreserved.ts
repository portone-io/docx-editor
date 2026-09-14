/**
 * How a block the editor never modelled goes back out, wherever in the document it stands.
 *
 * A block opened under the body names the fragment it stands for and the session holds the bytes;
 * one opened inside a table cell was never a fragment of its own and carries its XML along. An
 * edit may move either into the other's place, so the body writer and the cell writer read both
 * the same way, from here.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  DocxExportError,
  type ExportProblemReason,
  type ExportProblemStory,
} from "../ooxml/errors";
import type { ExportRefs } from "./exportRefs";
import { originalBlock, type SessionStore, splitBlockKey } from "./session";

/** Why there is no original to write, which a block that came in from another document answers differently */
export function lostOriginal(
  node: PMNode,
  session: SessionStore | null,
  story: ExportProblemStory | null
): { readonly message: string; readonly reason: ExportProblemReason } {
  const srcId: unknown = node.attrs.srcId;
  const key = typeof srcId === "string" ? splitBlockKey(srcId) : null;
  const name = node.type.name;
  if (key === null || key.sessionId === session?.sessionId) {
    return {
      message: "a preserved block has lost its original XML",
      reason: { kind: "lost-preserved-xml", node: name, story },
    };
  }
  return {
    message: `a preserved block comes from another document (${key.sessionId})`,
    reason: {
      kind: "preserved-from-another-document",
      node: name,
      sessionId: key.sessionId,
      story,
    },
  };
}

export function serializePreservedBlock(
  node: PMNode,
  refs: ExportRefs
): string {
  const xml: unknown = node.attrs.xml;
  if (typeof xml === "string") return xml;
  const imported = refs.session ? originalBlock(node, refs.session) : undefined;
  if (!imported) {
    // The writer is handed one block at a time and does not know which story it came out of, so
    // the reason names none; `docx/invariants` predicts the same refusal with the story on it
    const { message, reason } = lostOriginal(node, refs.session, null);
    throw new DocxExportError("lost-original", message, {
      problem: { code: "lost-original", message, reason },
    });
  }
  return imported.xml;
}
