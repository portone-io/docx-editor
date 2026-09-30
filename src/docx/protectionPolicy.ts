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
import {
  type EditableComments,
  type EditingProtection,
  unattributedCommentAuthors,
} from "../schema/protection";
import { type DocxBytes, importDocx } from "./importDocx";
import { CONTENT_TYPES_PATH, sameBytes } from "./packageParts";
import {
  type Relationship,
  readRelationships,
  relsPathOf,
  resolveTarget,
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

export interface JudgedPackages {
  arrived: SessionStore;
  submitted: SessionStore;
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
  /**
   * Grammar alone: whether the entry is one this editor's writer could have put out, whoever it
   * belongs to. `original` is the entry the file that arrived held under the same key, and null
   * for one that appeared, because what the writer can put out depends on what it had to work
   * with: it writes a comment's body out of the blocks it models and passes the rest through from
   * the entry it read.
   */
  wellFormed(entry: Element, original: Element | null): boolean;
  /**
   * Whether the entry came back differing from the one that arrived in nothing but a change this
   * protection leaves to everyone. Such an entry is nobody's rewrite, so it is held neither to
   * the grammar this editor writes in nor to who owns it.
   */
  anyonesChange(entry: Element, original: Element): boolean;
  /**
   * Permission alone: whether `authorId` may have written (`original === null`) or rewritten this
   * entry under `options`. `packages` lets a kind look across at a sibling part or back at what
   * arrived. `unattributed` is the display names the file that arrived writes comments under while
   * recording nobody for them, which is what an entry claiming no identity is held against.
   */
  allowed(
    entry: Element,
    original: Element | null,
    authorId: string,
    options: PolicyOptions,
    packages: JudgedPackages,
    unattributed: ReadonlySet<string>
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

/** The paths of the policy's parts this package holds, which is what a comparison of its bytes leaves to the policy */
export function policyPartPaths(
  policy: ProtectionPolicy<string, string>,
  session: SessionStore
): string[] {
  return partPathsOf(policy, session).filter(
    (path): path is string => path !== null
  );
}

/**
 * The main document part with every block of the body taken out of it: the namespaces the story
 * is written under, the text around the body, and the section properties that set the paper and
 * its margins. A comment is written inside a paragraph, so nothing a comment edit writes reaches
 * this text.
 *
 * The section closing the body is read off the document node rather than out of the tail
 * (`docx/sections`), and it belongs here for the same reason the rest does: a submission is free
 * to rewrite the paper it is written on, and a comparison of the blocks would not see it.
 */
export function aroundTheBlocks(story: Story): string {
  const sectPr: unknown = story.doc.attrs.sectPr;
  return (
    story.session.documentPrefix +
    (typeof sectPr === "string" ? sectPr : "") +
    story.session.documentSuffix
  );
}

/**
 * The blocks kept as the XML they arrived as, which the story comparison cannot see and so has to
 * be compared beside `aroundTheBlocks`.
 */
function preservedBlocks(story: Story): string {
  return story.session.blocks
    .filter((block) => !isModelledBlock(block.node))
    .map((block) => block.xml)
    .join("");
}

/**
 * Which of the policy's declarations two packages differ by: `gained` lets the revised package add
 * them, the way a submission writing its first comment does, and `gained or dropped` also lets it
 * give them up, the way a file that lost its last comment may.
 */
export type DeclarationChange = "gained" | "gained or dropped";

function partKindAt(
  policy: ProtectionPolicy<string, string>,
  session: SessionStore,
  path: string,
  matches: (kind: StoryPartKind) => boolean
): boolean {
  return partPathsOf(policy, session).some(
    (at, index) => at === path && matches(policy.parts[index])
  );
}

/**
 * Whether the relationships of the main document part are the ones it arrived with, save for the
 * policy's own parts. An id already handed out keeps pointing where it pointed.
 *
 * A part the file did not have may be gained, once, and only where the revised package holds it.
 * Writing under one of these protections relates each part a single time, so a second one under
 * the same type is not something this editor writes, and it is how a submission would otherwise
 * name a part of its choosing. A relationship may be dropped only where it is allowed to be and
 * the revised package no longer holds the part it named.
 *
 * An id names one relationship. A part naming one twice is read differently depending on which of
 * the two a reader keeps, so it is turned down rather than judged.
 */
function relationshipsKept(
  policy: ProtectionPolicy<string, string>,
  before: SessionStore,
  after: SessionStore,
  allowed: DeclarationChange
): boolean {
  const relsPath = relsPathOf(before.mainPartPath);
  const was = readRelationships(before.parts, relsPath);
  const now = readRelationships(after.parts, relsPath);
  if (
    new Set(was.map((entry) => entry.id)).size !== was.length ||
    new Set(now.map((entry) => entry.id)).size !== now.length
  ) {
    return false;
  }
  const relTypes = policy.parts.map((kind) => kind.relType);
  const heldByRevised = (entry: Relationship) =>
    after.parts.has(resolveTarget(before.mainPartPath, entry.target));
  const current = new Map(now.map((entry) => [entry.id, entry]));
  const kept = was.every((entry) => {
    const other = current.get(entry.id);
    if (other === undefined) {
      return (
        allowed === "gained or dropped" &&
        relTypes.includes(entry.type) &&
        !entry.external &&
        !heldByRevised(entry)
      );
    }
    return (
      other.type === entry.type &&
      other.target === entry.target &&
      other.external === entry.external
    );
  });
  const ids = new Set(was.map((entry) => entry.id));
  // Only the relationships a reader opens count: one pointing outside the package names no part,
  // whatever type it carries
  const parts = now.filter((entry) => !entry.external);
  return (
    kept &&
    now
      .filter((entry) => !ids.has(entry.id))
      .every(
        (entry) =>
          relTypes.includes(entry.type) &&
          !entry.external &&
          heldByRevised(entry) &&
          parts.filter((other) => other.type === entry.type).length === 1
      ) &&
    gainedPartsAreNew(policy, before, after)
  );
}

interface ContentTypes {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

/** What `[Content_Types].xml` declares: the types by extension, and the types of named parts */
function contentTypes(bytes: Uint8Array | undefined): ContentTypes {
  const declared: ContentTypes = { defaults: new Map(), overrides: new Map() };
  if (bytes === undefined) return declared;
  for (const el of elementChildren(
    parseXml(decodeUtf8(bytes).text).documentElement
  )) {
    const type = el.getAttribute("ContentType") ?? "";
    const extension = el.getAttribute("Extension");
    const partName = el.getAttribute("PartName");
    if (el.localName === "Default" && extension !== null) {
      declared.defaults.set(extension, type);
    } else if (el.localName === "Override" && partName !== null) {
      declared.overrides.set(partName, type);
    }
  }
  return declared;
}

function sameDeclarations(
  before: ReadonlyMap<string, string>,
  after: ReadonlyMap<string, string>
): boolean {
  return (
    before.size === after.size &&
    Array.from(before).every(([key, type]) => after.get(key) === type)
  );
}

/**
 * Whether the package declares the content types it arrived with, save for the override of a
 * policy part. A declaration that was there keeps naming the type it named.
 *
 * A gained override has to name a part of the policy where the revised package holds it, as the
 * type that part is: a type alone would excuse retyping any part, and a declaration by extension
 * types every part that ends in it. A dropped one has to name a policy part the original held and
 * the revised package no longer does.
 */
function contentTypesKept(
  policy: ProtectionPolicy<string, string>,
  before: SessionStore,
  after: SessionStore,
  allowed: DeclarationChange
): boolean {
  const was = contentTypes(before.parts.get(CONTENT_TYPES_PATH));
  const now = contentTypes(after.parts.get(CONTENT_TYPES_PATH));
  const pathOf = (partName: string) => partName.replace(/^\//, "");
  const dropped = Array.from(was.overrides).filter(
    ([partName]) => !now.overrides.has(partName)
  );
  const gained = Array.from(now.overrides).filter(
    ([partName]) => !was.overrides.has(partName)
  );
  return (
    sameDeclarations(was.defaults, now.defaults) &&
    Array.from(was.overrides).every(
      ([partName, type]) =>
        !now.overrides.has(partName) || now.overrides.get(partName) === type
    ) &&
    dropped.every(
      ([partName, type]) =>
        allowed === "gained or dropped" &&
        !after.parts.has(pathOf(partName)) &&
        partKindAt(
          policy,
          before,
          pathOf(partName),
          (kind) => kind.contentType === type
        )
    ) &&
    gained.every(
      ([partName, type]) =>
        after.parts.has(pathOf(partName)) &&
        partKindAt(
          policy,
          after,
          pathOf(partName),
          (kind) => kind.contentType === type
        )
    )
  );
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

/**
 * Whether each of the package's two declaration parts, the main part's relationships and
 * `[Content_Types].xml`, differs between the two packages in nothing but what the policy's parts
 * need, keyed by the part's path.
 */
export function declarationsKept(
  policy: ProtectionPolicy<string, string>,
  before: SessionStore,
  after: SessionStore,
  allowed: DeclarationChange
): ReadonlyMap<string, boolean> {
  return new Map([
    [
      relsPathOf(before.mainPartPath),
      relationshipsKept(policy, before, after, allowed),
    ],
    [CONTENT_TYPES_PATH, contentTypesKept(policy, before, after, allowed)],
  ]);
}

/** Whether every part outside the document story is the one the file arrived with */
function packageKept(
  policy: ProtectionPolicy<string, string>,
  before: Story,
  after: Story
): ChangeVerdict<never, PackageReason> {
  const was = before.session;
  const now = after.session;
  if (was.mainPartPath !== now.mainPartPath) {
    return { ok: false, reason: "part-changed", part: was.mainPartPath };
  }
  const relsPath = relsPathOf(was.mainPartPath);
  const untouched = new Set([
    was.mainPartPath,
    relsPath,
    CONTENT_TYPES_PATH,
    ...policyPartPaths(policy, was),
    ...policyPartPaths(policy, now),
  ]);
  for (const path of new Set([...was.parts.keys(), ...now.parts.keys()])) {
    if (untouched.has(path)) continue;
    const arrived = was.parts.get(path);
    const submitted = now.parts.get(path);
    if (
      arrived === undefined ||
      submitted === undefined ||
      !sameBytes(arrived, submitted)
    ) {
      return { ok: false, reason: "part-changed", part: path };
    }
  }
  const declarations = declarationsKept(policy, was, now, "gained");
  if (declarations.get(relsPath) !== true) {
    return { ok: false, reason: "relationship-changed", part: relsPath };
  }
  if (declarations.get(CONTENT_TYPES_PATH) !== true) {
    return { ok: false, reason: "part-changed", part: CONTENT_TYPES_PATH };
  }
  if (
    preservedBlocks(before) !== preservedBlocks(after) ||
    aroundTheBlocks(before) !== aroundTheBlocks(after)
  ) {
    return { ok: false, reason: "part-changed", part: was.mainPartPath };
  }
  return { ok: true };
}

/**
 * Whether every entry of one part came back as it was, or as one this editor writes for an author
 * who could have written it.
 *
 * An entry the file no longer stands behind is one no edit through the editor could have reached,
 * and the writer drops an entry only with what it stood for. Both halves hold for every part of
 * the policy: a comment nothing refers to, thread state or a date for no comment, an identity for
 * a name nobody writes under.
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

  const unattributed = unattributedCommentAuthors(
    before.session.comments.ordered
  );
  const stoodBehindNow = kind.referents(after);
  for (const [id, entry] of submitted) {
    const original = arrived.get(id);
    if (original && original.xml === entry.xml) continue;
    if (!stoodBehindNow.has(id)) return false;
    if (original && kind.anyonesChange(entry.el, original.el)) continue;
    if (
      !kind.wellFormed(entry.el, original?.el ?? null) ||
      !kind.allowed(
        entry.el,
        original?.el ?? null,
        authorId,
        options,
        { arrived: before.session, submitted: after.session },
        unattributed
      )
    ) {
      return false;
    }
  }

  const stoodBehindBefore = kind.referents(before);
  return Array.from(arrived.keys()).every(
    (id) =>
      submitted.has(id) ||
      (stoodBehindBefore.has(id) && !stoodBehindNow.has(id))
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
    const packaged = packageKept(policy, before, after);
    if (!packaged.ok) return packaged;
    const story = policy.storyKept(before, after, authorId, options);
    return story.ok
      ? partsKept(policy, before, after, authorId, options)
      : story;
  });
}
