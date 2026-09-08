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
import { NAMESPACES } from "../../ooxml/names";
import {
  attributeByLocalName,
  isElement,
  parseXml,
  serializeXml,
  W_NS,
} from "../../ooxml/xml";
import { isPreservedNode } from "../../schema/preservedFragments";
import {
  commentAdditionAllowed,
  type EditableComments,
} from "../../schema/protection";
import { NO_FORMATTING } from "../formatting";
import { NO_IMPORT_SOURCES } from "../importParagraph";
import { availablePartPath } from "../packageParts";
import type {
  EntryReading,
  StoryEntry,
  StoryPartKind,
} from "../protectionPolicy";
import { serializeBlock } from "../serializeBlock";
import type { SessionStore } from "../session";
import { buildBlock } from "../story";
import { isModelledBlock, type Story } from "../storyProjection";
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
  declarationsWritten,
  declarationWritten,
  holdsElementsAndLayout,
  isLayoutText,
  lastBodyParagraph,
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

/** Whether the node, or anything inside it, is content this editor keeps rather than models */
function holdsPreserved(node: PMNode): boolean {
  if (isPreservedNode(node)) return true;
  let found = false;
  node.descendants((child) => {
    if (isPreservedNode(child)) found = true;
    return !found;
  });
  return found;
}

/** The declarations a block is read under while it stands on its own, away from its part */
const BLOCK_DECLARATIONS = Object.entries(NAMESPACES)
  .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
  .join(" ");

/**
 * One block as this package reads and writes it, so that two spellings of it compare as one.
 *
 * Both sides go through the same parse under the same declarations: the entry's block as the
 * submission wrote it, and the same block as this writer puts it out. null for text that is no
 * readable XML at all.
 */
function normalizedBlock(xml: string): string | null {
  try {
    const root = parseXml(
      `<b ${BLOCK_DECLARATIONS}>${xml}</b>`
    ).documentElement;
    const [block] = Array.from(root.children);
    return block === undefined || root.children.length !== 1
      ? null
      : serializeXml(block);
  } catch {
    return null;
  }
}

/**
 * Whether this editor's writer could have written this block into a body from nothing.
 *
 * A body is a story now (`docx/story`), so the writer puts out whatever the editor models -
 * paragraphs and tables, their styles, their run formatting - and not one plain run any more. Two
 * things it never writes: content it does not model, a field or an `mc:AlternateContent` among it,
 * which it can only pass through from the entry that arrived; and bytes it did not read, since a
 * block it writes is one it can read back to exactly what it wrote. That fixed point is the same
 * one the document story is compared by (`docx/storyProjection`).
 */
function writtenBlock(el: Element): boolean {
  if (!declarationsWritten(el)) return false;
  const node = buildBlock(el, "", NO_IMPORT_SOURCES, NO_FORMATTING);
  if (!isModelledBlock(node) || holdsPreserved(node)) return false;
  const written = normalizedBlock(serializeBlock(node));
  return written !== null && written === normalizedBlock(serializeXml(el));
}

/**
 * Whether the body is one this editor's writer could have put out against the entry that arrived.
 *
 * Every block is either the writer's own output or a block of the original passed through
 * untouched, which is exactly what `serializeStory` writes: an edit of one paragraph leaves the
 * others as their bytes, the whitespace a producer laid them out over included (`./grammar`).
 * Nothing else stands between the blocks: anything a reader would show is a place to put bytes no
 * block comparison looks at.
 */
function wellFormedStoryBody(
  entry: Element,
  original: Element | null
): boolean {
  if (!holdsElementsAndLayout(entry)) return false;
  const kept = new Set(
    original === null ? [] : Array.from(original.children, serializeXml)
  );
  return Array.from(entry.children).every(
    (block) => kept.has(serializeXml(block)) || writtenBlock(block)
  );
}

/**
 * Whether this editor's writer could have put the entry out, whoever it belongs to.
 *
 * Shape alone: which attributes it carries and what stands inside it, judged against the entry the
 * original file held under the same id. Who may have written it is `entryAllowed`'s question.
 * Each kind is read where it is written (`./grammar`).
 */
export function wellFormedEntry(
  entry: Element,
  original: Element | null
): boolean {
  if (entry.namespaceURI === W_NS && entry.localName === "comment") {
    return (
      attributesWithin(entry, COMMENT_ATTRIBUTES) &&
      wellFormedStoryBody(entry, original)
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
 * An entry that appeared has to be one this author could have written, which
 * `commentAdditionAllowed` decides here as it does over the story. One that arrived keeps the
 * identity it arrived with; settling its thread or replying to it belong to everyone and leave
 * what it says alone, while rewriting what it says is its author's, a moderator's, or anyone's
 * where no identity was recorded for it. Both are the rules the editor holds to
 * (`schema/protection`), and the two have to answer alike or a file the editor wrote would be
 * turned down here.
 *
 * `unattributed` is read from the file that arrived, since it says what the submission is judged
 * against; `people` is the submission's, since it says who the submission names.
 */
export function entryAllowed(
  entry: Element,
  original: Element | null,
  authorId: string,
  editableComments: EditableComments,
  people: ImportedPeople,
  unattributed: ReadonlySet<string>
): boolean {
  if (entry.namespaceURI === W_NS && entry.localName === "comment") {
    const author = attributeByLocalName(entry, "author");
    const recorded = author === null ? null : commentAuthorId(people, author);
    if (original === null) {
      return commentAdditionAllowed(
        { author, authorId: recorded },
        authorId,
        unattributed
      );
    }
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
    const author = attributeByLocalName(entry, "author");
    return (
      original === null &&
      recordedIdentity(entry) === authorId &&
      author !== null &&
      !unattributed.has(author)
    );
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
  /** The element the entries stand in, which shares their namespace */
  rootName: string;
  /** Whether rebuilding the entries drops the content between them, as the two comment writers do */
  rewritesContents: boolean;
  localName: string;
  idAttr: string;
  xmlIn(session: SessionStore): string | null;
  pathIn(session: SessionStore): string | null;
  referents(story: Story): ReadonlySet<string>;
}

/** The compatibility declarations the writer adds to a part it writes a thread key into */
const COMPATIBILITY: ReadonlyMap<string, string> = new Map([
  ["xmlns:w14", NAMESPACES.w14],
  ["xmlns:mc", NAMESPACES.mc],
]);

/** The flag those declarations go with, which names the markup a reader may pass over */
const IGNORABLE = "mc:Ignorable";

/** The one thing the writer adds to it */
const IGNORED = "w14";

/** Content an entry comparison cannot see, inside or outside the part root. */
function surroundingNodes(parent: Node | null): readonly Node[] {
  return parent === null
    ? []
    : Array.from(parent.childNodes).filter(
        (node) => !isElement(node) && !isLayoutText(node)
      );
}

function sameNodes(now: readonly Node[], was: readonly Node[]): boolean {
  return (
    now.length === was.length &&
    now.every((node, at) => {
      const original = was[at];
      return (
        node.nodeType === original.nodeType &&
        node.nodeName === original.nodeName &&
        node.nodeValue === original.nodeValue
      );
    })
  );
}

/**
 * People are appended without replacing their surroundings. The comment writers rebuild their
 * entries and drop inter-entry content, but only an actual entry change can explain that loss.
 * In either case a submission cannot introduce or rewrite content outside the entries.
 */
function interEntryContentKept(
  now: Element,
  was: Element | null,
  rewritesContents: boolean
): boolean {
  const current = surroundingNodes(now);
  if (sameNodes(current, surroundingNodes(was))) return true;
  return (
    rewritesContents &&
    current.length === 0 &&
    was !== null &&
    Array.from(now.children, serializeXml).join("") !==
      Array.from(was.children, serializeXml).join("")
  );
}

function ignorableTokens(value: string): ReadonlySet<string> {
  return new Set(value.split(/\s+/).filter(Boolean));
}

/** Whether the compatibility flag still names what it named, and at most `w14` on top of it */
function ignorableKept(now: string, was: string | null): boolean {
  const gained = ignorableTokens(now);
  const had = was === null ? new Set<string>() : ignorableTokens(was);
  return (
    Array.from(had).every((token) => gained.has(token)) &&
    Array.from(gained).every((token) => had.has(token) || token === IGNORED)
  );
}

/** Whether the root is one this editor writes a part from nothing as */
function writtenRoot(root: Element, shape: CommentPartShape): boolean {
  if (root.namespaceURI !== shape.namespace) return false;
  if (root.localName !== shape.rootName) return false;
  return Array.from(root.attributes).every((attr) =>
    attr.name === IGNORABLE
      ? ignorableKept(attr.value, null)
      : declarationWritten(attr) || COMPATIBILITY.get(attr.name) === attr.value
  );
}

/**
 * Whether the element the entries stand in came back as it left.
 *
 * The writer keeps the opening tag it read and adds to it only what a thread key needs to mean
 * anything, so the root's name, the prefixes it binds and what it binds them to all have to come
 * back as they went. A part the package did not have is held to what this editor writes a new one
 * as instead: nothing else could have put it there.
 */
function rootKeptAgainst(
  now: Element,
  was: Element | null,
  shape: CommentPartShape
): boolean {
  if (was === null) return writtenRoot(now, shape);
  if (now.nodeName !== was.nodeName) return false;
  const had = new Map(
    Array.from(was.attributes, (attr) => [attr.name, attr.value] as const)
  );
  const kept = Array.from(now.attributes).every((attr) => {
    const before = had.get(attr.name) ?? null;
    if (attr.name === IGNORABLE) return ignorableKept(attr.value, before);
    if (before === null) return COMPATIBILITY.get(attr.name) === attr.value;
    return before === attr.value;
  });
  return kept && Array.from(had.keys()).every((name) => now.hasAttribute(name));
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
  for (const node of Array.from(parseXml(xml).documentElement.childNodes)) {
    // The root comparison checks non-entry content against the original, including annotations
    // the writer preserves. This reader only judges the entries it can key.
    if (!isElement(node)) continue;
    const el = node;
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

function storyPart(shape: CommentPartShape): StoryPartKind {
  return {
    relType: shape.relType,
    contentType: shape.contentType,
    pathIn: shape.pathIn,
    writePathIn: (session) =>
      shape.pathIn(session) ??
      availablePartPath(session.parts, session.mainPartPath, shape.baseName),
    entriesIn: (session, reading) =>
      entriesOf(shape, shape.xmlIn(session), reading),
    rootKept: (before, after) => {
      const submitted = shape.xmlIn(after);
      if (submitted === null) return true;
      const arrived = shape.xmlIn(before);
      const now = parseXml(submitted);
      const was = arrived === null ? null : parseXml(arrived);
      return (
        rootKeptAgainst(
          now.documentElement,
          was?.documentElement ?? null,
          shape
        ) &&
        sameNodes(surroundingNodes(now), surroundingNodes(was)) &&
        interEntryContentKept(
          now.documentElement,
          was?.documentElement ?? null,
          shape.rewritesContents
        )
      );
    },
    referents: shape.referents,
    wellFormed: wellFormedEntry,
    anyonesChange: threadKeyAlone,
    allowed: (entry, original, authorId, options, session, unattributed) =>
      entryAllowed(
        entry,
        original,
        authorId,
        options.editableComments,
        session.comments.people,
        unattributed
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
  rootName: "comments",
  rewritesContents: true,
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
  rootName: "commentsEx",
  rewritesContents: true,
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
  rootName: "people",
  rewritesContents: false,
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
