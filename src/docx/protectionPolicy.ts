/**
 * What a protection level lets a change rewrite, as one value every side of the package reads.
 *
 * A level says three things: which parts of the package a change may write, what an entry in one
 * of those parts may look like, and how the story reads once the markup the level is about is
 * taken out of it. The part planners write those parts and the verifier judges them, and each
 * carried its own copy of the first two with nothing but a test holding them together. Declaring
 * them once means a part the writer starts adding is a part the verifier already knows.
 *
 * A level with no policy registered is one no file is judged under: `protectionPolicyFor` answers
 * undefined, and a caller holding the level alone has nothing to run.
 */

import {
  decodeUtf8,
  elementChildren,
  parseXml,
  withXmlParser,
  type XmlParser,
} from "../ooxml/xml";
import type { EditableComments, EditingProtection } from "../schema/protection";
import { CONTENT_TYPES_PATH } from "./comments/constants";
import { type DocxBytes, importDocx } from "./importDocx";
import {
  type Relationship,
  readRelationships,
  relsPathOf,
} from "./relationships";
import type { SessionStore } from "./session";
import { isModelledBlock, type Story } from "./storyProjection";

/** One entry of a story part: the element read out of it, and its text for the byte-for-byte case */
export interface StoryEntry {
  el: Element;
  xml: string;
}

/**
 * How a part's entries are read.
 *
 * `arrived` is the reference a submission is held against rather than something to judge, so it
 * takes what it can read and passes over the rest, leaving an entry standing on something
 * unreadable to be judged as one that appeared. `submitted` is what is judged.
 */
export type EntryReading = "arrived" | "submitted";

/** What a judgement may be told about the reader beyond the two documents */
export interface PolicyOptions {
  editableComments: EditableComments;
}

/** What a judgement over two files may be told beyond that */
export interface VerifyOptions extends PolicyOptions {
  xmlParser?: XmlParser;
}

/** One package part a protection lets an editor rewrite, and how a rewritten entry is judged */
export interface StoryPartKind {
  relType: string;
  contentType: string;
  /** Where the reader found this part, and null for a package that holds none */
  pathIn(session: SessionStore): string | null;
  /**
   * Where the part goes when it is written: where it already sits, or the first name in the
   * story's own folder that no part of the package has taken.
   */
  writePathIn(session: SessionStore): string;
  /**
   * The part's entries keyed by the id the story or a sibling part refers to them by, and null
   * for a submitted part this editor's writer could not have put out at all.
   */
  entriesIn(
    session: SessionStore,
    reading: EntryReading
  ): ReadonlyMap<string, StoryEntry> | null;
  /**
   * Whether the element the entries stand in came back as it left, save for the compatibility
   * markup the writer declares on a part it writes a thread key into.
   */
  rootKept(before: SessionStore, after: SessionStore): boolean;
  /** The keys this file still stands behind, which is what an entry has to be keyed by */
  referents(story: Story): ReadonlySet<string>;
  /** Grammar alone: whether the entry is one this editor's writer could have put out, whoever it belongs to */
  wellFormed(entry: Element): boolean;
  /**
   * Whether the entry came back differing from the one that arrived in nothing but a change this
   * protection leaves to everyone. Such an entry is nobody's rewrite, so it is held neither to
   * the grammar this editor writes in nor to who owns it.
   */
  anyonesChange(entry: Element, original: Element): boolean;
  /**
   * Permission alone: whether `authorId` may have written (`original === null`) or rewritten this
   * entry under `options`. `session` is the submission's, so a kind can look across at a sibling
   * part, the way a comment's author resolves through the people part.
   */
  allowed(
    entry: Element,
    original: Element | null,
    authorId: string,
    options: PolicyOptions,
    session: SessionStore
  ): boolean;
}

/** The two reasons the package comparison answers with, whatever the policy */
export type PackageReason = "part-changed" | "relationship-changed";

/**
 * The shape every verdict takes: a story reason carries no part, a part reason names the part it
 * was reached over. A public verdict keeps its own declaration and is one instance of this.
 */
export type ChangeVerdict<
  StoryReason extends string,
  PartReason extends string,
> =
  | { ok: true }
  | { ok: false; reason: StoryReason }
  | { ok: false; reason: PartReason; part: string };

export interface ProtectionPolicy<
  StoryReason extends string,
  PartReason extends string,
> {
  level: EditingProtection;
  parts: readonly StoryPartKind[];
  /** The reason a rewritten entry is refused under when no part takes it as well formed and allowed */
  rejectedMarkup: PartReason;
  /** Whether the story reads the same once this protection's own markup is taken out */
  storyKept(
    before: Story,
    after: Story,
    authorId: string,
    options: PolicyOptions
  ): ChangeVerdict<StoryReason, PartReason>;
}

/**
 * The policies a run has loaded, under the level each judges.
 *
 * A policy registers itself as its own module is evaluated, so a level answers here only once
 * something has imported that module. `verifyChange` takes the policy itself and needs none of
 * this; the lookup is for a caller holding a level and nothing else.
 */
const policies = new Map<EditingProtection, ProtectionPolicy<string, string>>();

/** Records the policy under its level and hands it back, so a module declares and registers at once */
export function registerProtectionPolicy<S extends string, P extends string>(
  policy: ProtectionPolicy<S, P>
): ProtectionPolicy<S, P> {
  policies.set(policy.level, policy);
  return policy;
}

/**
 * The policy for a level, and undefined for a level no loaded module declares one for, `none` and
 * `readOnly` among them.
 */
export function protectionPolicyFor(
  level: EditingProtection
): ProtectionPolicy<string, string> | undefined {
  return policies.get(level);
}

function sameBytes(before: Uint8Array, after: Uint8Array): boolean {
  return (
    before.length === after.length &&
    before.every((byte, index) => byte === after[index])
  );
}

/**
 * Where the reader found each of the policy's parts, in the order the policy names them.
 *
 * These are the parts the byte comparison excuses, and they are the ones the reader opened rather
 * than every part a relationship of the right type points at. A package is free to relate a
 * second part under one of these types, and the reader takes the first of each; excusing the rest
 * would let a submission name any part it liked and have it go uncompared.
 */
function partPathsOf(
  policy: ProtectionPolicy<string, string>,
  session: SessionStore
): readonly (string | null)[] {
  return policy.parts.map((kind) => kind.pathIn(session));
}

function excusedPaths(
  policy: ProtectionPolicy<string, string>,
  session: SessionStore
): string[] {
  return partPathsOf(policy, session).filter(
    (path): path is string => path !== null
  );
}

/**
 * The main document part with the blocks the document model carries taken out of it.
 *
 * What is left is everything the document comparison cannot see: the namespaces the story is
 * written under, the section properties that set the paper and its margins, and every block kept
 * as the XML it arrived as. A comment is written inside a paragraph, so nothing a comment edit
 * writes reaches this text.
 */
function aroundTheStory(session: SessionStore): string {
  const preserved = session.blocks
    .filter((block) => !isModelledBlock(block.node))
    .map((block) => block.xml)
    .join("");
  return session.documentPrefix + preserved + session.documentSuffix;
}

/**
 * Whether the relationships of the main document part are the ones it arrived with, save for the
 * policy's own parts it may have gained. An id already handed out keeps pointing where it pointed.
 *
 * A part the file did not have may be gained, once. Writing under one of these protections relates
 * each part a single time, so a second one under the same type is not something this editor
 * writes, and it is how a submission would otherwise name a part of its choosing.
 *
 * An id names one relationship. A part naming one twice is read differently depending on which of
 * the two a reader keeps, so it is turned down rather than judged.
 */
function relationshipsKept(
  before: readonly Relationship[],
  after: readonly Relationship[],
  relTypes: readonly string[]
): boolean {
  if (
    new Set(before.map((entry) => entry.id)).size !== before.length ||
    new Set(after.map((entry) => entry.id)).size !== after.length
  ) {
    return false;
  }
  const now = new Map(after.map((entry) => [entry.id, entry]));
  const kept = before.every((entry) => {
    const current = now.get(entry.id);
    return (
      current !== undefined &&
      current.type === entry.type &&
      current.target === entry.target &&
      current.external === entry.external
    );
  });
  const ids = new Set(before.map((entry) => entry.id));
  // Every original relationship survives where `kept` holds, so a type standing once in the
  // submission is a type gained where the file had none. Only the relationships a reader opens
  // count: one pointing outside the package names no part, whatever type it carries
  const parts = after.filter((entry) => !entry.external);
  return (
    kept &&
    after
      .filter((entry) => !ids.has(entry.id))
      .every(
        (entry) =>
          relTypes.includes(entry.type) &&
          !entry.external &&
          parts.filter((other) => other.type === entry.type).length === 1
      )
  );
}

/** What `[Content_Types].xml` declares, each declaration under the name it is keyed by */
function contentTypes(bytes: Uint8Array | undefined): Map<string, string> {
  if (bytes === undefined) return new Map();
  const declared = new Map<string, string>();
  for (const el of elementChildren(
    parseXml(decodeUtf8(bytes).text).documentElement
  )) {
    const key =
      el.localName === "Default"
        ? el.getAttribute("Extension")
        : el.localName === "Override"
          ? el.getAttribute("PartName")
          : null;
    if (key !== null) declared.set(key, el.getAttribute("ContentType") ?? "");
  }
  return declared;
}

/**
 * Whether the package declares the content types it arrived with, save for an override a part it
 * gained needs. A declaration that was there keeps naming the type it named.
 */
function contentTypesKept(
  before: Map<string, string>,
  after: Map<string, string>,
  declarable: readonly string[]
): boolean {
  for (const [key, type] of before) {
    if (after.get(key) !== type) return false;
  }
  for (const [key, type] of after) {
    if (!before.has(key) && !declarable.includes(type)) return false;
  }
  return true;
}

/**
 * Whether a part the submission relates for the first time is one it brought with it.
 *
 * Writing the first comment writes a new part. Relating a part the file already had turns this
 * judgement's excuse for the policy's parts into an excuse for that part, whatever it holds.
 */
function gainedPartsAreNew(
  policy: ProtectionPolicy<string, string>,
  before: SessionStore,
  after: SessionStore
): boolean {
  const had = partPathsOf(policy, before);
  return partPathsOf(policy, after).every(
    (path, kind) =>
      path === null || path === had[kind] || !before.parts.has(path)
  );
}

/** Whether every part outside the document story is the one the file arrived with */
function packageKept(
  policy: ProtectionPolicy<string, string>,
  before: SessionStore,
  after: SessionStore
): ChangeVerdict<never, PackageReason> {
  if (before.mainPartPath !== after.mainPartPath) {
    return { ok: false, reason: "part-changed", part: before.mainPartPath };
  }
  const relsPath = relsPathOf(before.mainPartPath);
  const untouched = new Set([
    before.mainPartPath,
    relsPath,
    CONTENT_TYPES_PATH,
    ...excusedPaths(policy, before),
    ...excusedPaths(policy, after),
  ]);
  for (const path of new Set([...before.parts.keys(), ...after.parts.keys()])) {
    if (untouched.has(path)) continue;
    const was = before.parts.get(path);
    const now = after.parts.get(path);
    if (was === undefined || now === undefined || !sameBytes(was, now)) {
      return { ok: false, reason: "part-changed", part: path };
    }
  }
  if (
    !relationshipsKept(
      readRelationships(before.parts, relsPath),
      readRelationships(after.parts, relsPath),
      policy.parts.map((kind) => kind.relType)
    ) ||
    !gainedPartsAreNew(policy, before, after)
  ) {
    return { ok: false, reason: "relationship-changed", part: relsPath };
  }
  if (
    !contentTypesKept(
      contentTypes(before.parts.get(CONTENT_TYPES_PATH)),
      contentTypes(after.parts.get(CONTENT_TYPES_PATH)),
      policy.parts.map((kind) => kind.contentType)
    )
  ) {
    return { ok: false, reason: "part-changed", part: CONTENT_TYPES_PATH };
  }
  if (aroundTheStory(before) !== aroundTheStory(after)) {
    return { ok: false, reason: "part-changed", part: before.mainPartPath };
  }
  return { ok: true };
}

/**
 * Whether every entry of one part came back as it was, or as one this editor writes for an author
 * who could have written it.
 *
 * An entry the file no longer stands behind is one no edit through the editor could have reached,
 * and an entry it did not stand behind when it left is one no edit could have taken away. Both
 * halves hold for every part of the policy: a comment nothing refers to, thread state for no
 * comment, an identity for a name nobody writes under.
 */
function partKept(
  kind: StoryPartKind,
  before: Story,
  after: Story,
  authorId: string,
  options: PolicyOptions
): boolean {
  const arrived = kind.entriesIn(before.session, "arrived") ?? new Map();
  const submitted = kind.entriesIn(after.session, "submitted");
  if (submitted === null) return false;
  if (!kind.rootKept(before.session, after.session)) return false;

  const stoodBehindNow = kind.referents(after);
  for (const [id, entry] of submitted) {
    const original = arrived.get(id);
    if (original && original.xml === entry.xml) continue;
    if (!stoodBehindNow.has(id)) return false;
    if (original && kind.anyonesChange(entry.el, original.el)) continue;
    if (
      !kind.wellFormed(entry.el) ||
      !kind.allowed(
        entry.el,
        original?.el ?? null,
        authorId,
        options,
        after.session
      )
    ) {
      return false;
    }
  }

  const stoodBehindBefore = kind.referents(before);
  return Array.from(arrived.keys()).every(
    (id) => stoodBehindBefore.has(id) || submitted.has(id)
  );
}

/**
 * Whether the parts this protection lets an edit rewrite came back holding only entries this
 * editor writes, each of them one this author was in a position to write.
 *
 * This runs after the story has been judged, so a comment rewritten, re-anchored or deleted by the
 * wrong hand is already named for what it is. What is left to this is everything the story cannot
 * show: markup forged into a body, an entry nothing refers to, an identity recorded for somebody
 * else.
 */
export function partsKept<P extends string>(
  policy: ProtectionPolicy<string, P>,
  before: Story,
  after: Story,
  authorId: string,
  options: PolicyOptions
): ChangeVerdict<never, P> {
  for (const kind of policy.parts) {
    const path = kind.pathIn(after.session) ?? kind.pathIn(before.session);
    if (path === null) continue;
    if (!partKept(kind, before, after, authorId, options)) {
      return { ok: false, reason: policy.rejectedMarkup, part: path };
    }
  }
  return { ok: true };
}

/**
 * Whether the submitted file differs from the original in nothing this protection forbids.
 *
 * Both files are opened inside the one parser scope, so the parser named here is the parser both
 * reads go through even though neither `importDocx` call is given it. The package is judged first,
 * then the story, then the parts the policy excused from the package comparison: the parts are
 * last so that an edit the wrong hand made is named for the edit rather than for the part it was
 * written across.
 */
export function verifyChange<S extends string, P extends string>(
  policy: ProtectionPolicy<S, P>,
  original: DocxBytes,
  submitted: DocxBytes,
  authorId: string,
  { xmlParser, ...options }: VerifyOptions
): ChangeVerdict<S, P | PackageReason> {
  return withXmlParser(xmlParser, () => {
    const before = importDocx(original);
    const after = importDocx(submitted);
    const packaged = packageKept(policy, before.session, after.session);
    if (!packaged.ok) return packaged;
    const story = policy.storyKept(before, after, authorId, options);
    return story.ok
      ? partsKept(policy, before, after, authorId, options)
      : story;
  });
}
