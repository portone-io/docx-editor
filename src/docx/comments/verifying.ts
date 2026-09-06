/**
 * Whether every entry of the three comment parts is one this editor wrote, for an author who was
 * in a position to write it.
 *
 * The parts a comment is written across are the ones a submission may rewrite, so the package
 * comparison (`../commentOnlyChange`) leaves their bytes alone. Something has to read them, or a
 * file could carry anything at all inside a comment and still be answered as unchanged. What that
 * reading holds an entry to is two separate questions. Grammar asks whether this editor's writer
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
import type { CommentOnlyVerdict } from "../commentOnlyChange";
import type { Story } from "../storyProjection";
import { W14_NS, W15_NS } from "./constants";
import {
  attributesWithin,
  COMMENT_ATTRIBUTES,
  readStrictCommentBody,
  recordedIdentity,
  wellFormedCommentExtension,
  wellFormedPerson,
} from "./grammar";
import { commentReferencesIn } from "./model";
import { commentAuthorId, type ImportedPeople } from "./people";
import { lastBodyParagraph, lastParagraphId } from "./reading";

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

/** An entry as it is judged: the element to read, and the text two files are compared by */
interface Entry {
  el: Element;
  xml: string;
}

interface KeyedEntry extends Entry {
  id: string;
}

/**
 * The entries a part arrived with, which are the reference a submitted entry is held against
 * rather than something to judge. Anything unreadable is passed over, leaving the entry standing
 * on it to be judged as one that appeared.
 */
function arrivedEntries(
  xml: string | null,
  localName: string,
  idAttr: string
): ReadonlyMap<string, Entry> {
  const entries = new Map<string, Entry>();
  if (xml === null) return entries;
  for (const el of elementChildren(parseXml(xml).documentElement)) {
    if (el.localName !== localName) continue;
    const id = attributeByLocalName(el, idAttr);
    if (id === null || entries.has(id)) continue;
    entries.set(id, { el, xml: serializeXml(el) });
  }
  return entries;
}

/**
 * The entries a submitted part holds, and null for a part holding anything else.
 *
 * This editor writes one kind of child into each of these parts, one entry per id. A part carrying
 * a second kind of element, or two entries under one id, is not one it wrote, and reading it
 * leniently would leave whichever entry lost the key unjudged.
 */
function submittedEntries(
  xml: string | null,
  namespace: string,
  localName: string,
  idAttr: string
): KeyedEntry[] | null {
  if (xml === null) return [];
  const entries: KeyedEntry[] = [];
  const seen = new Set<string>();
  for (const el of elementChildren(parseXml(xml).documentElement)) {
    if (el.namespaceURI !== namespace || el.localName !== localName)
      return null;
    const id = attributeByLocalName(el, idAttr);
    if (id === null || seen.has(id)) return null;
    seen.add(id);
    entries.push({ id, el, xml: serializeXml(el) });
  }
  return entries;
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
function threadKeyAlone(entry: Element, original: Element): boolean {
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

/** The entries of a file's comments part, which is what its other two parts are written against */
function commentEntries(story: Story): readonly Element[] {
  return Array.from(
    arrivedEntries(story.session.comments.xml, "comment", "id").values(),
    (entry) => entry.el
  );
}

/**
 * One of the three parts a comment is written across: where it sits, what it holds, and what the
 * file has to still stand behind for an entry to belong there.
 */
interface PartKind {
  namespace: string;
  localName: string;
  idAttr: string;
  xmlOf(story: Story): string | null;
  pathOf(story: Story): string | null;
  /** The keys this file still stands behind, which is what an entry has to be keyed by */
  referents(story: Story): ReadonlySet<string>;
}

const COMMENT_PARTS: readonly PartKind[] = [
  {
    namespace: W_NS,
    localName: "comment",
    idAttr: "id",
    xmlOf: (story) => story.session.comments.xml,
    pathOf: (story) => story.session.comments.partPath,
    referents: (story) => referencedCommentIds(story.doc),
  },
  {
    namespace: W15_NS,
    localName: "commentEx",
    idAttr: "paraId",
    xmlOf: (story) => story.session.comments.extendedXml,
    pathOf: (story) => story.session.comments.extendedPartPath,
    // Thread state stands for a comment, and it is keyed by the comment's own thread key
    referents: (story) =>
      new Set(
        commentEntries(story).flatMap((el) => {
          const key = lastParagraphId(el);
          return key === null ? [] : [key];
        })
      ),
  },
  {
    namespace: W15_NS,
    localName: "person",
    idAttr: "author",
    xmlOf: (story) => story.session.comments.people.xml,
    pathOf: (story) => story.session.comments.people.partPath,
    // An identity stands for a name somebody writes comments under
    referents: (story) =>
      new Set(
        commentEntries(story).flatMap((el) => {
          const author = attributeByLocalName(el, "author");
          return author === null ? [] : [author];
        })
      ),
  },
];

/**
 * Whether every entry of one part came back as it was, or as one this editor writes for an author
 * who could have written it.
 *
 * An entry the file no longer stands behind is one no edit through the editor could have reached,
 * and an entry it did not stand behind when it left is one no edit could have taken away. Both
 * halves hold for each of the three parts: a comment nothing refers to, thread state for no
 * comment, an identity for a name nobody writes under.
 */
function partKept(
  kind: PartKind,
  before: Story,
  after: Story,
  authorId: string,
  editableComments: EditableComments
): boolean {
  const arrived = arrivedEntries(
    kind.xmlOf(before),
    kind.localName,
    kind.idAttr
  );
  const submitted = submittedEntries(
    kind.xmlOf(after),
    kind.namespace,
    kind.localName,
    kind.idAttr
  );
  if (submitted === null) return false;

  const stoodBehindNow = kind.referents(after);
  const people = after.session.comments.people;
  for (const entry of submitted) {
    const original = arrived.get(entry.id);
    if (original && original.xml === entry.xml) continue;
    if (!stoodBehindNow.has(entry.id)) return false;
    if (original && threadKeyAlone(entry.el, original.el)) continue;
    if (
      !wellFormedEntry(entry.el) ||
      !entryAllowed(
        entry.el,
        original?.el ?? null,
        authorId,
        editableComments,
        people
      )
    ) {
      return false;
    }
  }

  const held = new Set(submitted.map((entry) => entry.id));
  const stoodBehindBefore = kind.referents(before);
  return Array.from(arrived.keys()).every(
    (id) => stoodBehindBefore.has(id) || held.has(id)
  );
}

/**
 * Whether the three comment parts came back holding only entries this editor writes, each of them
 * one this author was in a position to write.
 *
 * This runs after the story has been judged, so a comment rewritten, re-anchored or deleted by the
 * wrong hand is already named for what it is (`../commentOnlyChange`). What is left to this is
 * everything the story cannot show: markup forged into a body, an entry nothing refers to, an
 * identity recorded for somebody else.
 */
export function commentPartsKept(
  before: Story,
  after: Story,
  authorId: string,
  editableComments: EditableComments
): CommentOnlyVerdict {
  for (const kind of COMMENT_PARTS) {
    const path = kind.pathOf(after) ?? kind.pathOf(before);
    if (path === null) continue;
    if (!partKept(kind, before, after, authorId, editableComments)) {
      return { ok: false, reason: "part-changed", part: path };
    }
  }
  return { ok: true };
}
