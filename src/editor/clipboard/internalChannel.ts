/**
 * The channel a copy made in this editor comes back through.
 *
 * What leaves for the clipboard is HTML, and HTML carries only what a reader can make out of it:
 * a paragraph's own properties, a table's grid, the relationship a link hung off are none of them
 * writable there without publishing the document's insides. So the slice itself stays here, under
 * a name the copied HTML carries, and a paste that hands that name back is given what was copied.
 *
 * The name is trusted only alongside the session it was copied from: another open document -
 * a second editor on the same page among them - numbers its lists and relates its images against
 * its own file, so a slice of it means something else here, and such a paste falls to the reader
 * that reads the markup instead. A token nothing kept, a token from a copy since replaced, and a
 * token invented by whoever wrote the markup all miss the same way.
 *
 * The slice is kept alongside the kind of each list its paragraphs name, which is the one thing
 * about the copy that only the editor it was copied from can answer: a number says nothing about
 * whether it counted or bulleted, and the definition that said so is the source document's. It is
 * read here because this is where both ends are the same session.
 */

import type { Node as PMNode, Slice } from "prosemirror-model";
import type { ListKind } from "../../numbering/listTemplate";
import type { Numbering } from "../../numbering/parseNumbering";
import { docxSchema } from "../../schema";
import { listKindOf, listRefOf } from "../commands/listCommands";

/** Where the copied HTML carries the name of the slice this editor kept */
export const INTERNAL_TOKEN_ATTRIBUTE = "data-docx-clip";

/** The kind of list each numbering id of a slice stood for where the slice was copied */
export type ListKinds = ReadonlyMap<number, ListKind>;

export const NO_LIST_KINDS: ListKinds = new Map();

interface Copied {
  readonly token: string;
  readonly sessionId: string;
  readonly slice: Slice;
  readonly listKinds: ListKinds;
}

/** What one name stands for: the slice that was copied, and what its list numbers meant there */
export interface RecalledCopy {
  readonly slice: Slice;
  readonly listKinds: ListKinds;
}

/**
 * The last copy made in this process, which is the only one a paste can name.
 *
 * A clipboard holds one thing at a time, so keeping more would be keeping slices nothing can ask
 * for. The entry lives until the next copy replaces it or the process ends; a paste that arrives
 * after a copy made anywhere else - another editor, another document - finds a token that is no
 * longer the one it carries.
 */
let copied: Copied | null = null;

let copies = 0;

/**
 * A name for one copy. It is drawn fresh every time rather than worked out from what was copied,
 * so a paste of an older copy of the very same content cannot answer to the name of a newer one.
 */
function nextToken(): string {
  copies += 1;
  return `c${copies}-${Math.random().toString(36).slice(2, 10)}`;
}

/**
 * Keeps the copied slice under a fresh name, and answers with the name to write into the markup.
 * Null for a state built without a document: with no session to bind the name to there is nothing
 * to tell this editor's copy from another's, and the copy travels as markup alone.
 */
export function rememberCopied(
  slice: Slice,
  sessionId: string | null,
  numbering: Numbering
): string | null {
  if (sessionId === null) return null;
  const token = nextToken();
  copied = {
    token,
    sessionId,
    slice,
    listKinds: listKindsIn(slice, numbering),
  };
  return token;
}

/** The copy that name stands for, and only when the session asking is the one that copied it */
export function recallCopied(
  token: string | null,
  sessionId: string | null
): RecalledCopy | null {
  if (copied === null || token === null || sessionId === null) return null;
  return copied.token === token && copied.sessionId === sessionId
    ? { slice: copied.slice, listKinds: copied.listKinds }
    : null;
}

/** What the numbering of the source document says each list the slice names is */
function listKindsIn(slice: Slice, numbering: Numbering): ListKinds {
  const kinds = new Map<number, ListKind>();
  const read = (node: PMNode): void => {
    if (node.type === docxSchema.nodes.paragraph) {
      const ref = listRefOf(node);
      const kind = ref === null ? null : listKindOf(numbering, ref);
      if (ref !== null && kind !== null) kinds.set(ref.numId, kind);
      return;
    }
    node.forEach(read);
  };
  slice.content.forEach(read);
  return kinds.size === 0 ? NO_LIST_KINDS : kinds;
}

/**
 * The name the pasted markup carries, if it carries one.
 *
 * The copy writes it on the first element it draws. `prosemirror-view` may wrap that element in
 * the parents its tag needs (a row in a table) and descends back through them before parsing, so
 * the element handed here is either the one carrying the name or its parent.
 */
export function internalTokenOf(dom: Node): string | null {
  if (dom instanceof Element) {
    const own = dom.getAttribute(INTERNAL_TOKEN_ATTRIBUTE);
    if (own !== null) return own;
  }
  const first =
    dom instanceof Element || dom instanceof DocumentFragment
      ? dom.firstElementChild
      : null;
  return first?.getAttribute(INTERNAL_TOKEN_ATTRIBUTE) ?? null;
}
