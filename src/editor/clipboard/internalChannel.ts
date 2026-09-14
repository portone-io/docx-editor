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
 *
 * The note each reference in it calls is kept the same way, and for a stronger reason: a cut takes
 * a footnote away together with its last reference (`editor/plugins/noteLifecycle`), so by the time
 * the paste happens the document no longer holds what the copy was pointing at.
 */

import type { Node as PMNode, Slice } from "prosemirror-model";
import type { ListKind } from "../../numbering/listTemplate";
import type { Numbering } from "../../numbering/parseNumbering";
import { docxSchema } from "../../schema";
import {
  EDITABLE_NOTE_KINDS,
  type NoteKey,
  type StoryKey,
  storyKey,
  storyNodeOf,
} from "../../schema/stories";
import { listKindOf, listRefOf } from "../commands/listCommands";

/** Where the copied HTML carries the name of the slice this editor kept */
export const INTERNAL_TOKEN_ATTRIBUTE = "data-docx-clip";

/** The kind of list each numbering id of a slice stood for where the slice was copied */
export type ListKinds = ReadonlyMap<number, ListKind>;

export const NO_LIST_KINDS: ListKinds = new Map();

/** The story of each note the slice refers to, as the document copied from held it */
export type NoteStories = ReadonlyMap<NoteKey, PMNode>;

export const NO_NOTE_STORIES: NoteStories = new Map();

/**
 * The two ways a copy leaves the editor.
 *
 * A drag is serialized through the very same path a copy takes, so the one the drag wrote would
 * replace the one the clipboard still holds, and a drag begun and let go would send the next paste
 * of that clipboard to the reader that reads markup. They are kept apart instead, and each is
 * replaced only by the next copy made the same way.
 */
export type CopyRoute = "clipboard" | "drag";

interface Copied {
  readonly token: string;
  readonly sessionId: string;
  readonly slice: Slice;
  readonly listKinds: ListKinds;
  readonly noteStories: NoteStories;
}

/** What one name stands for: the slice that was copied, and what it named where it was copied */
export interface RecalledCopy {
  readonly slice: Slice;
  readonly listKinds: ListKinds;
  readonly noteStories: NoteStories;
}

/** The document a copy is made from, which is the only one that can answer what it carries */
export interface CopySource {
  /** Null for a state built without a document, which has no session to bind a name to */
  readonly sessionId: string | null;
  readonly doc: PMNode;
  readonly numbering: Numbering;
  /** The notes laying a part out rather than saying something, which are never copied */
  readonly specialNotes: ReadonlySet<StoryKey>;
}

/**
 * The last copy made each way in this process, which are the only ones a paste can name.
 *
 * A clipboard holds one thing at a time and so does a drag, so keeping more would be keeping
 * slices nothing can ask for. An entry lives until the next copy made that way replaces it or the
 * process ends; a paste that arrives after a copy made anywhere else - another editor, another
 * document - finds a token that is no longer one of the two it carries.
 */
const copied = new Map<CopyRoute, Copied>();

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
  route: CopyRoute,
  slice: Slice,
  source: CopySource
): string | null {
  const { sessionId } = source;
  if (sessionId === null) return null;
  const token = nextToken();
  copied.set(route, {
    token,
    sessionId,
    slice,
    listKinds: listKindsIn(slice, source.numbering),
    noteStories: noteStoriesIn(slice, source),
  });
  return token;
}

function recalled(copy: Copied): RecalledCopy {
  return {
    slice: copy.slice,
    listKinds: copy.listKinds,
    noteStories: copy.noteStories,
  };
}

/** The copy that name stands for, and only when the session asking is the one that copied it */
export function recallCopied(
  token: string | null,
  sessionId: string | null
): RecalledCopy | null {
  if (token === null || sessionId === null) return null;
  const found = [...copied.values()].find(
    (copy) => copy.token === token && copy.sessionId === sessionId
  );
  return found === undefined ? null : recalled(found);
}

/**
 * The notes the last drag out of this editor carried.
 *
 * A drop is handed the slice itself rather than markup (`prosemirror-view`), so it carries no name
 * to look a copy up by; the drag is the one route whose last copy is the copy being dropped. The
 * session must still be the one that dragged, since a drop from another editor arrives as the
 * markup its own reader answers for.
 */
export function draggedNotes(sessionId: string | null): NoteStories {
  const found = copied.get("drag");
  return found === undefined ||
    sessionId === null ||
    found.sessionId !== sessionId
    ? NO_NOTE_STORIES
    : found.noteStories;
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
 * The story of every note the slice calls, which is what a paste of it has to put back.
 *
 * Only the kinds an edit may add are kept: a reference to any other kind is dropped by the paste
 * (`./normalizers`), and so is one calling a note that lays a part out rather than saying
 * something, or one the file arrived with no note behind.
 */
function noteStoriesIn(slice: Slice, source: CopySource): NoteStories {
  const stories = new Map<NoteKey, PMNode>();
  const read = (node: PMNode): void => {
    if (node.type === docxSchema.nodes.noteReference) {
      const kind = EDITABLE_NOTE_KINDS.find(
        (candidate) => candidate === node.attrs.kind
      );
      const id: unknown = node.attrs.id;
      if (kind === undefined || typeof id !== "string") return;
      const key = storyKey(kind, id);
      const story = source.specialNotes.has(key)
        ? null
        : storyNodeOf(source.doc, key);
      if (story !== null) stories.set(key, story);
      return;
    }
    node.forEach(read);
  };
  slice.content.forEach(read);
  return stories.size === 0 ? NO_NOTE_STORIES : stories;
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
