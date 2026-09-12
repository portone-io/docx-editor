/**
 * The side stories a document holds beside its body: a comment's text, a footnote's, a header's.
 *
 * Each is a document of this same schema, held on the document node under the key that names it,
 * so an edit to one rides a transaction and lands in the history the way a body edit does. What
 * turns one into XML and back is `docx/story`; here is only what the schema layer needs, which is
 * the key, the value on the node, and the two comparisons the protection rules are written over.
 */

import type { Node as PMNode } from "prosemirror-model";
import { sameSource } from "./sourceEquality";

/** Which of the two notes parts a note stands in */
export const NOTE_KINDS = ["footnote", "endnote"] as const;

export type NoteKind = (typeof NOTE_KINDS)[number];

/** The two kinds of story a header or footer part holds, one story to a part */
export const HEADER_FOOTER_KINDS = ["header", "footer"] as const;

export type HeaderFooterKind = (typeof HEADER_FOOTER_KINDS)[number];

export const STORY_KINDS = [
  "comment",
  ...NOTE_KINDS,
  ...HEADER_FOOTER_KINDS,
] as const;

export type StoryKind = (typeof STORY_KINDS)[number];

/** `comment:4`, `footnote:2`, `header:word/header1.xml` */
export type StoryKey = `${StoryKind}:${string}`;

/**
 * What a reader listening to the page is told a note reference is.
 *
 * A reference whose own mark follows it in the document draws no label of the editor's
 * (§17.11.14), so it is named by its kind alone rather than by a number that is not there.
 */
export function noteName(kind: NoteKind, label: string): string {
  const named = kind === "endnote" ? "Endnote" : "Footnote";
  return label === "" ? named : `${named} ${label}`;
}

/** One story as the document node holds it, which is `Node.toJSON` of a document of this schema */
export interface StoryJson {
  readonly type: string;
  readonly attrs?: Readonly<Record<string, unknown>>;
  readonly content?: readonly unknown[];
}

/** The attr the stories stand on, named here so the schema and the readers cannot spell it apart */
export const STORIES_ATTR = "stories";

export function storyKey<K extends StoryKind>(
  kind: K,
  id: string
): `${K}:${string}` {
  return `${kind}:${id}`;
}

/** The key of a story that is a note: `footnote:2`, `endnote:3` */
export type NoteKey = `${NoteKind}:${string}`;

/**
 * The note this key names, and null for a key naming a story of another kind.
 *
 * A note's id never carries a colon - a note is identified by a whole number (§17.11.2, §17.11.8)
 * - so the kind is everything up to the first one.
 */
export function noteKeyOf(
  key: StoryKey
): { readonly kind: NoteKind; readonly id: string } | null {
  const at = key.indexOf(":");
  const named = key.slice(0, at);
  const kind = NOTE_KINDS.find((candidate) => candidate === named);
  return kind === undefined ? null : { kind, id: key.slice(at + 1) };
}

/**
 * The kinds of note an edit may add, delete, copy and rewrite, which is both of them: each part is
 * written back one entry per note (`docx/notes/writing`). The lifecycle, the clipboard and the
 * navigation all read this one list, so a kind of note added to the format joins it once.
 */
export const EDITABLE_NOTE_KINDS: readonly NoteKind[] = NOTE_KINDS;

/**
 * The key this text spells, or null for one that names no kind of story.
 *
 * An id carries colons of its own - a header is named by its part path - so the kind is read off
 * the front and everything after the first colon is the id.
 */
export function asStoryKey(text: string): StoryKey | null {
  const at = text.indexOf(":");
  if (at < 1) return null;
  const named = text.slice(0, at);
  const kind = STORY_KINDS.find((candidate) => candidate === named);
  return kind === undefined ? null : storyKey(kind, text.slice(at + 1));
}

function isStoryJson(value: unknown): value is StoryJson {
  return (
    typeof value === "object" &&
    value !== null &&
    typeof Reflect.get(value, "type") === "string"
  );
}

/** The stories this document holds, under the key naming each */
export function storiesOf(doc: PMNode): Readonly<Record<string, StoryJson>> {
  const held: unknown = doc.attrs[STORIES_ATTR];
  if (typeof held !== "object" || held === null) return {};
  const stories: Record<string, StoryJson> = {};
  for (const [key, value] of Object.entries(held)) {
    if (asStoryKey(key) !== null && isStoryJson(value)) stories[key] = value;
  }
  return stories;
}

/**
 * `Node.fromJSON` once per JSON value.
 *
 * The value on the node is never rewritten in place - an edit sets a new one - so a cache keyed by
 * the object itself hands back the same node for as long as the story says the same thing, which
 * is what lets a reader ask for a story on every render without rebuilding it.
 */
const built = new WeakMap<StoryJson, PMNode>();

/** The story this key names as a document node, and null where the document holds none */
export function storyNodeOf(doc: PMNode, key: StoryKey): PMNode | null {
  const json = storiesOf(doc)[key];
  if (json === undefined) return null;
  const cached = built.get(json);
  if (cached) return cached;
  const node = doc.type.schema.nodeFromJSON(json);
  built.set(json, node);
  return node;
}

/** The document node with a block key on nothing, which is what two sessions can be compared by */
const unkeyed = new WeakMap<PMNode, PMNode>();

function withoutBlockKeys(node: PMNode): PMNode {
  if (node.isText) return node;
  const cached = unkeyed.get(node);
  if (cached) return cached;
  const children = node.children.map(withoutBlockKeys);
  const attrs =
    node.attrs.srcId === null || node.attrs.srcId === undefined
      ? node.attrs
      : { ...node.attrs, srcId: null };
  const stripped =
    attrs === node.attrs &&
    children.every((child, at) => child === node.child(at))
      ? node
      : node.type.create(attrs, children, node.marks);
  unkeyed.set(node, stripped);
  return stripped;
}

/**
 * Whether the two stories say the same thing.
 *
 * The block keys are left out of it, because they name the document each story was read from and
 * two files opened side by side never share one. Everything else a block would be written back as
 * is compared (`./sourceEquality`), so a body that gained a bold run reads as changed and one
 * whose display values were worked out again does not.
 */
export function sameStory(a: PMNode | null, b: PMNode | null): boolean {
  if (a === null || b === null) return a === b;
  return sameSource(withoutBlockKeys(a), withoutBlockKeys(b));
}

/**
 * The document with its comment stories taken out of it, replies included.
 *
 * A comment edit now moves the document node, and a comparison that could not see past that would
 * call every comment edit a body edit. The footnote and header stories stay: changing one of them
 * is a change to what the document says, and a comment protection is meant to turn it down.
 */
export function withoutCommentStories(doc: PMNode): PMNode {
  const stories = storiesOf(doc);
  const kept = Object.fromEntries(
    Object.entries(stories).filter(([key]) => !key.startsWith("comment:"))
  );
  if (Object.keys(kept).length === Object.keys(stories).length) return doc;
  return doc.type.create(
    { ...doc.attrs, [STORIES_ATTR]: kept },
    doc.content,
    doc.marks
  );
}
