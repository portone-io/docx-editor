/**
 * The side stories a part beside the body is written from, told apart by what an edit did to each.
 *
 * Whether a story has to be written again is one question every story writer asks, so it is asked
 * here once. A writer answering it for itself would be free to keep a story another writer drops,
 * and the export invariants would have a third answer of their own.
 */

import type { Node as PMNode } from "prosemirror-model";
import { attrsText, openTagXml, type XmlAttr } from "../ooxml/element";
import type { DocxExportErrorCode } from "../ooxml/errors";
import { NAMESPACES, wName } from "../ooxml/names";
import {
  ensureRootDeclarations,
  partRootProblem,
  type RootDeclarations,
  splicePart,
} from "../ooxml/partSplice";
import { parseAttrs, readTag, rootTagAt } from "../ooxml/tagScan";
import {
  attributeByLocalName,
  decodeUtf8,
  elementChildren,
  encodeUtf8,
  parseXml,
} from "../ooxml/xml";
import { sameSource } from "../schema/sourceEquality";
import {
  asStoryKey,
  type StoryKey,
  type StoryKind,
  storiesOf,
  storyKey,
  storyNodeOf,
} from "../schema/stories";
import { type ExportRefs, NO_EXPORT_REFS } from "./exportRefs";
import { withUniqueStoryIdentities } from "./identities";
import {
  availablePartPath,
  CONTENT_TYPES_PATH,
  readPart,
  relatedPartPath,
} from "./packageParts";
import type { PartPlanContext, PartPlanner } from "./partPlan";
import { directoryOf } from "./relationships";
import { scanBlocksIn } from "./scan";
import { type StoryContainer, serializeStory } from "./serializeStory";
import type { SessionStore } from "./session";
import type { ImportedStory } from "./story";

export type StoryChange =
  | { readonly change: "kept"; readonly imported: ImportedStory }
  | {
      readonly change: "edited";
      readonly imported: ImportedStory;
      readonly current: PMNode;
    }
  | { readonly change: "removed"; readonly imported: ImportedStory }
  | {
      readonly change: "added";
      readonly key: StoryKey;
      readonly current: PMNode;
    };

const DECIMAL = /^-?\d+$/;

/** Numbers in numeric order, which is the order an entry id is counted in; anything else by its text */
function byId(a: string, b: string): number {
  if (DECIMAL.test(a) && DECIMAL.test(b)) return Number(a) - Number(b);
  if (a === b) return 0;
  return a < b ? -1 : 1;
}

/** The id a key of this kind names */
export function storyIdOf(key: StoryKey, kind: StoryKind): string {
  return key.slice(kind.length + 1);
}

function arrivedChange(doc: PMNode, imported: ImportedStory): StoryChange {
  const current = storyNodeOf(doc, imported.key);
  if (current === null) return { change: "removed", imported };
  return sameSource(current, imported.doc)
    ? { change: "kept", imported }
    : { change: "edited", imported, current };
}

/** Every story of this kind the package arrived with or the document now holds, in part order then id order */
export function storyChangesOf(
  doc: PMNode,
  session: SessionStore,
  kind: StoryKind
): readonly StoryChange[] {
  const arrived = Array.from(session.stories.values())
    .filter((imported) => imported.kind === kind)
    .map((imported) => arrivedChange(doc, imported));
  const added = Object.keys(storiesOf(doc))
    .flatMap((text) => {
      const key = asStoryKey(text);
      return key === null ||
        !key.startsWith(`${kind}:`) ||
        session.stories.has(key)
        ? []
        : [key];
    })
    .sort((a, b) => byId(storyIdOf(a, kind), storyIdOf(b, kind)))
    .flatMap((key): StoryChange[] => {
      const current = storyNodeOf(doc, key);
      return current === null ? [] : [{ change: "added", key, current }];
    });
  return [...arrived, ...added];
}

/** Whether any story of the kind changed, which is whether a part holding them is written at all */
function changed(changes: readonly StoryChange[]): boolean {
  return changes.some((change) => change.change !== "kept");
}

/** The kinds of story some part writer carries into the file; a change to any other kind has nowhere to go */
export const WRITTEN_STORY_KINDS: ReadonlySet<StoryKind> = new Set<StoryKind>([
  "header",
  "footer",
  "comment",
  "footnote",
]);

/** A part holding one entry per story, e.g. `w:footnotes` holding a `w:footnote` apiece */
export interface StoryEntriesPart {
  readonly name: string;
  readonly kind: StoryKind;
  readonly relType: string;
  readonly contentType: string;
  /** The file name a part this planner creates takes beside the main part, ahead of any number */
  readonly stem: string;
  /** The local name of the part's root element */
  readonly root: string;
  /** The local name of the element each story stands in */
  readonly entry: string;
  /** The ids the document's references to this part name, which a part this planner creates keeps clear of */
  referencedIds(doc: PMNode): ReadonlySet<string>;
  /** The entries a part this planner creates opens with, before any story */
  prelude(taken: ReadonlySet<string>): string;
}

/** One story as its part is written with it, beside what an edit did to it */
interface WrittenStory {
  readonly change: StoryChange;
  readonly story: PMNode;
}

/**
 * The story an entry goes out as, and null for one the part leaves out.
 *
 * A special entry - a separator, a continuation notice - lays the page out rather than saying
 * anything, so it goes back out as it arrived whatever the document now says of it, and
 * `docx/invariants` refuses such a change before anything is written.
 */
function writtenStoryOf(
  change: StoryChange,
  special: ReadonlySet<StoryKey>
): PMNode | null {
  if (change.change === "added") return change.current;
  if (change.change === "kept" || special.has(change.imported.key)) {
    return change.imported.doc;
  }
  return change.change === "edited" ? change.current : null;
}

function writtenStories(
  changes: readonly StoryChange[],
  session: SessionStore
): readonly WrittenStory[] {
  return changes.flatMap((change): WrittenStory[] => {
    const story = writtenStoryOf(change, session.specialNotes);
    return story === null ? [] : [{ change, story }];
  });
}

/**
 * The stories a part of this kind is written with, in the order it writes them, and none where no
 * story of the kind changed and the part is not written at all. The identity pass runs over this
 * list as one part, and the export invariants ask the same list.
 */
export function storyEntriesOf(
  doc: PMNode,
  session: SessionStore,
  kind: StoryKind
): readonly PMNode[] {
  const changes = storyChangesOf(doc, session, kind);
  return changed(changes)
    ? writtenStories(changes, session).map(({ story }) => story)
    : [];
}

/** One reason the part cannot take the changes its stories went through */
export interface StoryPartProblem {
  readonly code: DocxExportErrorCode;
  readonly message: string;
}

/**
 * Why the part cannot take what the document did to its stories, or none when it can. The part the
 * package holds is rewritten around its root element, and a part the package lacks is declared in
 * the content types part, which the export does not write from nothing.
 */
export function storyEntriesProblems(
  part: StoryEntriesPart,
  doc: PMNode,
  session: SessionStore
): readonly StoryPartProblem[] {
  if (!changed(storyChangesOf(doc, session, part.kind))) return [];
  const xml = readPart(
    session.parts,
    relatedPartPath(session.parts, session.mainPartPath, part.relType)
  );
  if (xml !== null) {
    const problem = partRootProblem(xml, part.root);
    return problem === null
      ? []
      : [{ code: "malformed-xml", message: problem }];
  }
  return session.parts.has(CONTENT_TYPES_PATH)
    ? []
    : [
        {
          code: "missing-content-types",
          message: `cannot add a part to a package that has no ${CONTENT_TYPES_PATH}`,
        },
      ];
}

/** What a rewritten or created part declares: every entry this writer puts out is spelled under `w` */
const ENTRY_MARKUP: RootDeclarations = { namespaces: { w: NAMESPACES.w } };

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function containerOf(part: StoryEntriesPart, id: string): StoryContainer {
  const name = wName(part.entry);
  return {
    open: openTagXml(name, attrsText([[wName("id"), id]])),
    close: `</${name}>`,
  };
}

/** One entry as the part writes it: an untouched one as the bytes it arrived as, an edited one block by block */
function entryXml(
  part: StoryEntriesPart,
  { change }: WrittenStory,
  settled: PMNode,
  special: ReadonlySet<StoryKey>,
  refs: ExportRefs
): string {
  if (change.change === "added") {
    return serializeStory(
      settled,
      null,
      containerOf(part, storyIdOf(change.key, part.kind)),
      refs
    );
  }
  const { imported } = change;
  return special.has(imported.key) || settled === imported.doc
    ? imported.xml
    : serializeStory(settled, imported, imported, refs);
}

/** What stands between the last entry and the root's closing tag, which is the layout a producer left there */
function beforeClosingTag(suffix: string): string {
  let at = suffix.indexOf("<");
  while (at !== -1) {
    const tag = readTag(suffix, at);
    if (tag === null || tag.kind === "close") return suffix.slice(0, at);
    at = suffix.indexOf("<", tag.end);
  }
  return suffix;
}

/**
 * The part with its entries in the order it holds them, then every story added since.
 *
 * What stands in the root and is no story - an entry naming no id, anything else a producer put
 * there - goes back out as it arrived, and the second entry naming an id already written is left
 * out, since every reader takes the first (`docx/story`).
 */
function rewrittenPart(
  part: StoryEntriesPart,
  xml: string,
  arrived: ReadonlyMap<StoryKey, string>,
  appended: string
): string {
  const scan = scanBlocksIn(xml, (_tag, depth) => depth === 0);
  const seen = new Set<string>();
  const entries = elementChildren(parseXml(xml).documentElement).map(
    (el, at) => {
      const slice = scan?.blocks[at]?.xml ?? "";
      const id =
        el.localName === part.entry ? attributeByLocalName(el, "id") : null;
      if (id === null) return slice;
      if (seen.has(id)) return "";
      seen.add(id);
      const entry = arrived.get(storyKey(part.kind, id));
      return entry === undefined
        ? ""
        : slice.slice(0, slice.indexOf("<")) + entry;
    }
  );
  return ensureRootDeclarations(
    splicePart(xml, {
      root: part.root,
      replaceChildren:
        entries.join("") + appended + beforeClosingTag(scan?.suffix ?? ""),
    }),
    ENTRY_MARKUP
  );
}

/**
 * The namespaces the main part's root binds, and what it lets a reader ignore.
 *
 * A block moved out of the body keeps the markup it was written in there - a `w14:paraId`, a `w15`
 * property - so a part this writer creates binds every prefix the way the body does, which is also
 * how Word writes every part of a package.
 */
function mainDeclarations(documentPrefix: string): readonly XmlAttr[] {
  const at = rootTagAt(documentPrefix);
  const tag = at === -1 ? null : readTag(documentPrefix, at);
  if (tag === null || tag.kind !== "open") return [];
  const attrs = parseAttrs(documentPrefix.slice(tag.nameEnd, tag.end - 1));
  return (attrs ?? []).filter(
    ([name]) =>
      (name.startsWith("xmlns:") && name !== "xmlns:w") ||
      name === "mc:Ignorable"
  );
}

function createdPart(
  part: StoryEntriesPart,
  session: SessionStore,
  children: string
): string {
  const name = wName(part.root);
  const declarations = attrsText(mainDeclarations(session.documentPrefix));
  return ensureRootDeclarations(
    `${XML_DECLARATION}${openTagXml(name, declarations === "" ? null : declarations)}${children}</${name}>`,
    ENTRY_MARKUP
  );
}

function planEntries(
  part: StoryEntriesPart,
  doc: PMNode,
  session: SessionStore,
  context: PartPlanContext
): ReadonlyMap<string, Uint8Array> | null {
  const changes = storyChangesOf(doc, session, part.kind);
  if (!changed(changes)) return null;
  const refs: ExportRefs = {
    ...NO_EXPORT_REFS,
    notes: context.notes,
    session,
  };
  const written = writtenStories(changes, session);
  const settled = withUniqueStoryIdentities(written.map(({ story }) => story));
  const entries = written.map((entry, at) => ({
    change: entry.change,
    xml: entryXml(
      part,
      entry,
      settled[at] ?? entry.story,
      session.specialNotes,
      refs
    ),
  }));
  const appended = entries
    .flatMap(({ change, xml }) => (change.change === "added" ? [xml] : []))
    .join("");
  const related = relatedPartPath(
    session.parts,
    session.mainPartPath,
    part.relType
  );
  const path =
    related ??
    availablePartPath(session.parts, session.mainPartPath, part.stem);
  const bytes = session.parts.get(path);
  if (bytes !== undefined) {
    const { text, hadBom } = decodeUtf8(bytes);
    const arrived = new Map(
      entries.flatMap(({ change, xml }): [StoryKey, string][] =>
        change.change === "added" ? [] : [[change.imported.key, xml]]
      )
    );
    return new Map([
      [path, encodeUtf8(rewrittenPart(part, text, arrived, appended), hadBom)],
    ]);
  }
  if (related === null) {
    context.relationships.add({
      type: part.relType,
      target: path.slice(directoryOf(session.mainPartPath).length),
    });
  }
  context.contentTypes.addOverride(path, part.contentType);
  const taken = new Set([
    ...changes.map((change) =>
      change.change === "added"
        ? storyIdOf(change.key, part.kind)
        : change.imported.id
    ),
    ...part.referencedIds(doc),
  ]);
  return new Map([
    [
      path,
      encodeUtf8(
        createdPart(part, session, part.prelude(taken) + appended),
        false
      ),
    ],
  ]);
}

/**
 * A planner for a part holding one entry per story.
 *
 * Nothing is written while every story of the kind stands as it arrived, so an untouched package
 * hands the part back as its own bytes. Once one changed, the part is written in the order it holds
 * its entries: an untouched entry as it arrived, an edited one again, a removed one not at all,
 * and every story added since after them in id order. A package holding no such part gets one,
 * with its relationship, its content type, and the entries `prelude` opens it with.
 */
export function storyEntriesPlanner(part: StoryEntriesPart): PartPlanner {
  return {
    name: part.name,
    plan: (doc, session, context) => planEntries(part, doc, session, context),
  };
}
