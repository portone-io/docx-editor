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
import { elementChildren, parseXml, serializeXml, W_NS } from "../../ooxml/xml";
import type { EditableComments } from "../../schema/protection";
import type { CommentOnlyVerdict } from "../commentOnlyChange";
import type { Story } from "../storyProjection";
import {
  attributesWithin,
  COMMENT_ATTRIBUTES,
  COMMENT_EX_ATTRIBUTES,
  readStrictCommentBody,
} from "./bodyGrammar";
import { COMMENT_AUTHOR_PROVIDER, W14_NS, W15_NS } from "./constants";
import { commentReferencesIn } from "./model";
import { commentAuthorId, type ImportedPeople } from "./people";

const PERSON_ATTRIBUTES: ReadonlySet<string> = new Set([`${W15_NS} author`]);

const PRESENCE_ATTRIBUTES: ReadonlySet<string> = new Set([
  `${W15_NS} providerId`,
  `${W15_NS} userId`,
]);

/** The attributes of a comment that say whose it is, which nobody rewrites, its own author included */
const COMMENT_IDENTITY: readonly string[] = [
  "id",
  "author",
  "date",
  "initials",
];

function attribute(el: Element, localName: string): string | null {
  return (
    Array.from(el.attributes).find((entry) => entry.localName === localName)
      ?.value ?? null
  );
}

function sameAttributes(
  entry: Element,
  original: Element,
  names: readonly string[]
): boolean {
  return names.every(
    (name) => attribute(entry, name) === attribute(original, name)
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
    const id = attribute(el, idAttr);
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
    const id = attribute(el, idAttr);
    if (id === null || seen.has(id)) return null;
    seen.add(id);
    entries.push({ id, el, xml: serializeXml(el) });
  }
  return entries;
}

/** Every comment the story stands for, replies included, which is every entry it still refers to */
export function referencedCommentIds(doc: PMNode): ReadonlySet<string> {
  const ids = new Set<string>();
  for (const [id, comment] of commentReferencesIn(doc)) {
    ids.add(id);
    for (const reply of comment.replies) ids.add(reply.id);
  }
  return ids;
}

function personWellFormed(entry: Element): boolean {
  if (!attributesWithin(entry, PERSON_ATTRIBUTES)) return false;
  const children = elementChildren(entry);
  if (children.length !== 1) return false;
  const [presence] = children;
  return (
    presence.namespaceURI === W15_NS &&
    presence.localName === "presenceInfo" &&
    attributesWithin(presence, PRESENCE_ATTRIBUTES) &&
    attribute(presence, "providerId") === COMMENT_AUTHOR_PROVIDER &&
    attribute(presence, "userId") !== null &&
    elementChildren(presence).length === 0
  );
}

/**
 * Whether this editor's writer could have put the entry out, whoever it belongs to.
 *
 * Shape alone: which attributes it carries and what stands inside it. Who may have written it is
 * `entryAllowed`'s question.
 */
export function wellFormedEntry(entry: Element): boolean {
  if (entry.namespaceURI === W_NS && entry.localName === "comment") {
    return (
      attributesWithin(entry, COMMENT_ATTRIBUTES) &&
      readStrictCommentBody(entry) !== null
    );
  }
  if (entry.namespaceURI === W15_NS && entry.localName === "commentEx") {
    return (
      attributesWithin(entry, COMMENT_EX_ATTRIBUTES) &&
      elementChildren(entry).length === 0
    );
  }
  if (entry.namespaceURI === W15_NS && entry.localName === "person") {
    return personWellFormed(entry);
  }
  return false;
}

/**
 * The key a comment's thread state is written against, which the writer puts on the paragraph
 * rather than on the entry (`./bodyGrammar`). Read the way the importer reads it.
 */
function commentParaId(entry: Element): string | null {
  const paragraphs = Array.from(entry.getElementsByTagNameNS(W_NS, "p"));
  const last = paragraphs[paragraphs.length - 1];
  return last === undefined ? null : attribute(last, "paraId");
}

/**
 * The entry with the thread key taken off, which is what two files are compared by when the one
 * thing between them is that a thread was settled or replied to.
 *
 * The key is taken off and put back rather than copied away from: an entry carries prefixes the
 * part root declares, so the fragment on its own does not parse.
 */
function withoutThreadKey(entry: Element): string {
  const paragraphs = Array.from(entry.getElementsByTagNameNS(W_NS, "p"));
  const last = paragraphs[paragraphs.length - 1];
  const key = last?.getAttributeNodeNS(W14_NS, "paraId") ?? null;
  if (last === undefined || key === null) return serializeXml(entry);
  last.removeAttributeNode(key);
  const said = serializeXml(entry);
  last.setAttributeNodeNS(key);
  return said;
}

function recordedIdentity(person: Element): string | null {
  const [presence] = elementChildren(person);
  return presence === undefined ? null : attribute(presence, "userId");
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
    const author = attribute(entry, "author");
    const recorded = author === null ? null : commentAuthorId(people, author);
    if (original === null) return recorded === authorId;
    // A thread key appears the first time a comment is settled or replied to, but one already
    // written is what its thread state hangs off and is not re-pointed
    const paraId = commentParaId(original);
    if (
      !sameAttributes(entry, original, COMMENT_IDENTITY) ||
      (paraId !== null && commentParaId(entry) !== paraId)
    ) {
      return false;
    }
    if (withoutThreadKey(entry) === withoutThreadKey(original)) return true;
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

function judged(
  entry: KeyedEntry,
  arrived: ReadonlyMap<string, Entry>,
  authorId: string,
  editableComments: EditableComments,
  people: ImportedPeople
): boolean {
  const original = arrived.get(entry.id);
  if (original && original.xml === entry.xml) return true;
  return (
    wellFormedEntry(entry.el) &&
    entryAllowed(
      entry.el,
      original?.el ?? null,
      authorId,
      editableComments,
      people
    )
  );
}

function commentsKept(
  before: Story,
  after: Story,
  authorId: string,
  editableComments: EditableComments
): boolean {
  const arrived = arrivedEntries(before.session.comments.xml, "comment", "id");
  const submitted = submittedEntries(
    after.session.comments.xml,
    W_NS,
    "comment",
    "id"
  );
  if (submitted === null) return false;

  const referredToBefore = referencedCommentIds(before.doc);
  const referredToNow = referencedCommentIds(after.doc);
  const people = after.session.comments.people;

  for (const entry of submitted) {
    const original = arrived.get(entry.id);
    if (original && original.xml === entry.xml) continue;
    // Nothing in the story points at this entry, so no edit through the editor reached it
    if (!referredToNow.has(entry.id)) return false;
    if (!judged(entry, arrived, authorId, editableComments, people)) {
      return false;
    }
  }

  // An entry nothing referred to when the file left was not one an edit could have taken away
  const held = new Set(submitted.map((entry) => entry.id));
  return Array.from(arrived.keys()).every(
    (id) => referredToBefore.has(id) || held.has(id)
  );
}

function extensionsKept(
  before: Story,
  after: Story,
  authorId: string,
  editableComments: EditableComments
): boolean {
  const arrived = arrivedEntries(
    before.session.comments.extendedXml,
    "commentEx",
    "paraId"
  );
  const submitted = submittedEntries(
    after.session.comments.extendedXml,
    W15_NS,
    "commentEx",
    "paraId"
  );
  if (submitted === null) return false;

  const paragraphs = new Set(
    Array.from(
      arrivedEntries(after.session.comments.xml, "comment", "id").values()
    ).flatMap((entry) => {
      const paraId = commentParaId(entry.el);
      return paraId === null ? [] : [paraId];
    })
  );

  return submitted.every(
    (entry) =>
      // A thread state stands for a comment, so one that appeared has to name a comment there is
      (arrived.has(entry.id) || paragraphs.has(entry.id)) &&
      judged(
        entry,
        arrived,
        authorId,
        editableComments,
        after.session.comments.people
      )
  );
}

function peopleKept(
  before: Story,
  after: Story,
  authorId: string,
  editableComments: EditableComments
): boolean {
  const people = after.session.comments.people;
  const arrived = arrivedEntries(
    before.session.comments.people.xml,
    "person",
    "author"
  );
  const submitted = submittedEntries(people.xml, W15_NS, "person", "author");
  if (submitted === null) return false;
  return submitted.every((entry) =>
    judged(entry, arrived, authorId, editableComments, people)
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
  const commentsPath =
    after.session.comments.partPath ?? before.session.comments.partPath;
  if (
    commentsPath !== null &&
    !commentsKept(before, after, authorId, editableComments)
  ) {
    return { ok: false, reason: "part-changed", part: commentsPath };
  }

  const extendedPath =
    after.session.comments.extendedPartPath ??
    before.session.comments.extendedPartPath;
  if (
    extendedPath !== null &&
    !extensionsKept(before, after, authorId, editableComments)
  ) {
    return { ok: false, reason: "part-changed", part: extendedPath };
  }

  const peoplePath =
    after.session.comments.people.partPath ??
    before.session.comments.people.partPath;
  if (
    peoplePath !== null &&
    !peopleKept(before, after, authorId, editableComments)
  ) {
    return { ok: false, reason: "part-changed", part: peoplePath };
  }
  return { ok: true };
}
