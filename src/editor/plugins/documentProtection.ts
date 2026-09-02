/**
 * Holds the protection the editor runs under (`schema/protection`) in the state, where the guard
 * and the commands read it. The guard itself stands in `lockedContent`, which enforces
 * `transactionAllowed` for locks and protection in one pass.
 */

import { Plugin, type Transaction } from "prosemirror-state";
import type {
  EditableComments,
  EditingProtection,
  ProtectionState,
} from "../../schema/protection";
import { protectionKey } from "../../schema/protectionState";
import type { CommentAuthor } from "../commands/comments/model";

export interface DocumentProtectionOptions {
  protection: EditingProtection;
  /** Whose comments are being written. Null where none are, a reader above all */
  author: CommentAuthor | null;
  editableComments: EditableComments;
}

function rulesOf({
  protection,
  author,
  editableComments,
}: DocumentProtectionOptions): ProtectionState {
  return { protection, authorId: author?.id ?? null, editableComments };
}

function isRules(value: unknown): value is ProtectionState {
  if (typeof value !== "object" || value === null) return false;
  const protection: unknown = Reflect.get(value, "protection");
  const authorId: unknown = Reflect.get(value, "authorId");
  const editableComments: unknown = Reflect.get(value, "editableComments");
  return (
    typeof protection === "string" &&
    (authorId === null || typeof authorId === "string") &&
    typeof editableComments === "string"
  );
}

/**
 * The transaction that puts the editor under another protection, which is how a mode is switched
 * on a document already open: `reconfigure` keeps the state a plugin of the same key already
 * holds, so a new plugin instance would not do it.
 */
export function setProtection(
  tr: Transaction,
  options: DocumentProtectionOptions
): Transaction {
  return tr.setMeta(protectionKey, rulesOf(options));
}

export function documentProtection(
  options: DocumentProtectionOptions
): Plugin<ProtectionState> {
  const rules = rulesOf(options);
  return new Plugin({
    key: protectionKey,
    state: {
      init: () => rules,
      apply: (transaction, value) => {
        const next: unknown = transaction.getMeta(protectionKey);
        return isRules(next) ? next : value;
      },
    },
  });
}
