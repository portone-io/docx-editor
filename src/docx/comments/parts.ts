/**
 * The three parts a comment is written across, each declared once: where it sits, what it is
 * declared as in the package, what an entry in it looks like, and who may have written one.
 *
 * The writer named the relationship and the content type where it added a part and the verifier
 * named them again where it excused one, so the two were free to drift apart. Both read this
 * instead, which is what makes a part the editor writes a part it takes back.
 *
 * What an entry is held to is two separate questions. Grammar asks whether this editor's writer
 * could have put the entry out, whoever it belongs to. Permission asks whether this author could
 * have made the change. Asking them as one would refuse a moderator rewriting somebody else's
 * comment under `editableComments: "all"`, which is a thing the editor lets them do.
 */

import type { Node as PMNode } from "prosemirror-model";
import {
  attributeByLocalName,
  elementChildren,
  isElement,
  parseXml,
  serializeXml,
  W_NS,
} from "../../ooxml/xml";
import type { EditableComments } from "../../schema/protection";
import type {
  EntryReading,
  StoryEntry,
  StoryPartKind,
} from "../protectionPolicy";
import { directoryOf } from "../relationships";
import type { SessionStore } from "../session";
import type { Story } from "../storyProjection";
import {
  COMMENTS_CONTENT_TYPE,
  COMMENTS_EXTENDED_CONTENT_TYPE,
  COMMENTS_EXTENDED_REL_TYPE,
  COMMENTS_REL_TYPE,
  PEOPLE_CONTENT_TYPE,
  PEOPLE_REL_TYPE,
  W14_NS,
  W15_NS,
} from "./constants";
import {
  attributesWithin,
  COMMENT_ATTRIBUTES,
  lastBodyParagraph,
  readStrictCommentBody,
  recordedIdentity,
  wellFormedCommentExtension,
  wellFormedPerson,
} from "./grammar";
import { commentReferencesIn } from "./model";
import { commentAuthorId, type ImportedPeople } from "./people";
import { lastParagraphId } from "./reading";

/** The attributes of a comment that say whose it is, which nobody rewrites, its own author included */
const COMMENT_IDENTITY: readonly string[] = [
  "id",
  "author",
  "date",
  "initials",
];

function sameAttributes(
  entry: Element,
  original: Element,
  names: readonly string[]
): boolean {
  return names.every(
    (name) =>
      attributeByLocalName(entry, name) === attributeByLocalName(original, name)
  );
}

/**
 * Whether this editor's writer could have put the entry out, whoever it belongs to.
 *
 * Shape alone: which attributes it carries and what stands inside it. Who may have written it is
 * `entryAllowed`'s question. Each kind is read where it is written (`./grammar`).
 */
export function wellFormedEntry(entry: Element): boolean {
  if (entry.namespaceURI === W_NS && entry.localName === "comment") {
    return (
      attributesWithin(entry, COMMENT_ATTRIBUTES) &&
      readStrictCommentBody(entry) !== null
    );
  }
  if (entry.namespaceURI === W15_NS && entry.localName === "commentEx") {
    return wellFormedCommentExtension(entry);
  }
  if (entry.namespaceURI === W15_NS && entry.localName === "person") {
    return wellFormedPerson(entry);
  }
  return false;
}

/**
 * The entry with the thread key taken off, which is what two files are compared by when the one
 * thing between them is that a thread was settled or replied to.
 */
function withoutThreadKey(entry: Element): string {
  const copy = entry.cloneNode(true);
  if (!isElement(copy)) return serializeXml(entry);
  const last = lastBodyParagraph(copy);
  last?.removeAttributeNS(W14_NS, "paraId");
  return serializeXml(copy);
}

/**
 * Whether the only thing between the two is that a thread gained its key.
 *
 * Settling a thread or replying to it belongs to everyone and leaves what the comment says alone,
 * and the writer answers by putting the key on the entry that arrived rather than writing one of
 * its own (`./grammar`). Such an entry is not a rewrite, so it is neither held to the grammar this
 * editor writes bodies in nor to who owns the comment.
 */
export function threadKeyAlone(entry: Element, original: Element): boolean {
  const before = lastParagraphId(original);
  const after = lastParagraphId(entry);
  const kept = before === null || after === before;
  return kept && withoutThreadKey(entry) === withoutThreadKey(original);
}

/**
 * Whether `authorId` could have written or rewritten this entry.
 *
 * Permission alone: whether the shape is one this editor writes is `wellFormedEntry`'s question.
 * An entry that appeared has to be this author's. One that arrived keeps the identity it arrived
 * with; settling its thread or replying to it belong to everyone and leave what it says alone,
 * while rewriting what it says is its author's, a moderator's, or anyone's where no identity was
 * recorded for it. That last is the rule the editor holds to (`schema/protection`), and the two
 * have to answer alike or a file the editor wrote would be turned down here.
 */
export function entryAllowed(
  entry: Element,
  original: Element | null,
  authorId: string,
  editableComments: EditableComments,
  people: ImportedPeople
): boolean {
  if (entry.namespaceURI === W_NS && entry.localName === "comment") {
    const author = attributeByLocalName(entry, "author");
    const recorded = author === null ? null : commentAuthorId(people, author);
    if (original === null) return recorded === authorId;
    // A thread key appears the first time a comment is settled or replied to, but one already
    // written is what its thread state hangs off and is not re-pointed
    const paraId = lastParagraphId(original);
    if (
      !sameAttributes(entry, original, COMMENT_IDENTITY) ||
      (paraId !== null && lastParagraphId(entry) !== paraId)
    ) {
      return false;
    }
    return (
      editableComments === "all" || recorded === null || recorded === authorId
    );
  }
  if (entry.namespaceURI === W15_NS && entry.localName === "commentEx") {
    // Settling a comment and reopening it are everyone's, so `done` is the one thing that moves
    if (original === null) return true;
    return sameAttributes(entry, original, ["paraId", "paraIdParent"]);
  }
  if (entry.namespaceURI === W15_NS && entry.localName === "person") {
    // An identity already recorded is nobody's to rewrite, its own subject included
    return original === null && recordedIdentity(entry) === authorId;
  }
  return false;
}

/** What one of the three parts holds, beyond what every story part declares */
interface CommentPartShape {
  relType: string;
  contentType: string;
  /** The name a new part takes, ahead of the number that tells it from one the package has */
  baseName: string;
  namespace: string;
  localName: string;
  idAttr: string;
  xmlIn(session: SessionStore): string | null;
  pathIn(session: SessionStore): string | null;
  referents(story: Story): ReadonlySet<string>;
}

/**
 * The entries of a part, keyed by the id each carries.
 *
 * The lenient reading of what arrived passes over anything it cannot key, since it is only the
 * reference a submitted entry is held against. The strict reading of a submission answers null
 * instead: this editor writes one kind of child into each of these parts, one entry per id, so a
 * part carrying a second kind or two entries under one id is not one it wrote, and reading it
 * leniently would leave whichever entry lost the key unjudged.
 */
function entriesOf(
  shape: CommentPartShape,
  xml: string | null,
  reading: EntryReading
): ReadonlyMap<string, StoryEntry> | null {
  const entries = new Map<string, StoryEntry>();
  if (xml === null) return entries;
  const strict = reading === "submitted";
  for (const el of elementChildren(parseXml(xml).documentElement)) {
    const named = strict
      ? el.namespaceURI === shape.namespace && el.localName === shape.localName
      : el.localName === shape.localName;
    if (!named) {
      if (strict) return null;
      continue;
    }
    const id = attributeByLocalName(el, shape.idAttr);
    if (id === null || entries.has(id)) {
      if (strict) return null;
      continue;
    }
    entries.set(id, { el, xml: serializeXml(el) });
  }
  return entries;
}

/** The entries a package arrived holding, which the lenient reading never refuses */
function arrivedIn(
  shape: CommentPartShape,
  session: SessionStore
): ReadonlyMap<string, StoryEntry> {
  return entriesOf(shape, shape.xmlIn(session), "arrived") ?? new Map();
}

/** The first name in the story's own folder that no part of the package has taken */
function availablePath(session: SessionStore, baseName: string): string {
  const directory = directoryOf(session.mainPartPath);
  for (let suffix = 0; ; suffix += 1) {
    const path = `${directory}${baseName}${suffix === 0 ? "" : suffix + 1}.xml`;
    if (!session.parts.has(path)) return path;
  }
}

function storyPart(shape: CommentPartShape): StoryPartKind {
  return {
    relType: shape.relType,
    contentType: shape.contentType,
    pathIn: shape.pathIn,
    writePathIn: (session) =>
      shape.pathIn(session) ?? availablePath(session, shape.baseName),
    entriesIn: (session, reading) =>
      entriesOf(shape, shape.xmlIn(session), reading),
    referents: shape.referents,
    wellFormed: wellFormedEntry,
    anyonesChange: threadKeyAlone,
    allowed: (entry, original, authorId, options, session) =>
      entryAllowed(
        entry,
        original,
        authorId,
        options.editableComments,
        session.comments.people
      ),
  };
}

/** Every comment the story stands for, replies included, which is every entry it still refers to */
function referencedCommentIds(doc: PMNode): ReadonlySet<string> {
  const ids = new Set<string>();
  for (const [id, comment] of commentReferencesIn(doc)) {
    ids.add(id);
    for (const reply of comment.replies) ids.add(reply.id);
  }
  return ids;
}

const commentsShape: CommentPartShape = {
  relType: COMMENTS_REL_TYPE,
  contentType: COMMENTS_CONTENT_TYPE,
  baseName: "comments",
  namespace: W_NS,
  localName: "comment",
  idAttr: "id",
  xmlIn: (session) => session.comments.xml,
  pathIn: (session) => session.comments.partPath,
  referents: (story) => referencedCommentIds(story.doc),
};

/** The entries of a file's comments part, which is what its other two parts are written against */
function commentEntries(story: Story): readonly Element[] {
  return Array.from(
    arrivedIn(commentsShape, story.session).values(),
    (entry) => entry.el
  );
}

export const commentsPart: StoryPartKind = storyPart(commentsShape);

export const commentsExtendedPart: StoryPartKind = storyPart({
  relType: COMMENTS_EXTENDED_REL_TYPE,
  contentType: COMMENTS_EXTENDED_CONTENT_TYPE,
  baseName: "commentsExtended",
  namespace: W15_NS,
  localName: "commentEx",
  idAttr: "paraId",
  xmlIn: (session) => session.comments.extendedXml,
  pathIn: (session) => session.comments.extendedPartPath,
  // Thread state stands for a comment, and it is keyed by the comment's own thread key
  referents: (story) =>
    new Set(
      commentEntries(story).flatMap((el) => {
        const key = lastParagraphId(el);
        return key === null ? [] : [key];
      })
    ),
});

export const peoplePart: StoryPartKind = storyPart({
  relType: PEOPLE_REL_TYPE,
  contentType: PEOPLE_CONTENT_TYPE,
  baseName: "people",
  namespace: W15_NS,
  localName: "person",
  idAttr: "author",
  xmlIn: (session) => session.comments.people.xml,
  pathIn: (session) => session.comments.people.partPath,
  // An identity stands for a name somebody writes comments under
  referents: (story) =>
    new Set(
      commentEntries(story).flatMap((el) => {
        const author = attributeByLocalName(el, "author");
        return author === null ? [] : [author];
      })
    ),
});

/** The parts a comment protection lets an edit rewrite, in the order an export writes them */
export const COMMENT_STORY_PARTS: readonly [
  StoryPartKind,
  StoryPartKind,
  StoryPartKind,
] = [commentsPart, commentsExtendedPart, peoplePart];
