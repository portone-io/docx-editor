/**
 * The people part, where Word records who a comment author is beyond the display name
 * (`w15:people`, [MS-DOCX] §2.5.43 `presenceInfo`): one `w15:person` per author name, carrying a
 * `w15:presenceInfo` whose `w15:providerId` names the directory that issued `w15:userId`.
 *
 * This editor records the identity a host application hands it under a provider of its own
 * (`COMMENT_AUTHOR_PROVIDER`). An author name is the key on both sides, as it is for Word, so two
 * identities writing under one name cannot be told apart once the document is closed.
 */

import { DocxExportError } from "../../ooxml/errors";
import {
  decodeUtf8,
  elementChildren,
  encodeUtf8,
  escapeXml,
  parseXml,
  serializeXml,
} from "../../ooxml/xml";
import {
  directoryOf,
  type RelationshipWriter,
  readRelationships,
  relsPathOf,
  resolveTarget,
} from "../relationships";
import type { SessionStore } from "../session";
import {
  COMMENT_AUTHOR_PROVIDER,
  CONTENT_TYPES_PATH,
  PEOPLE_CONTENT_TYPE,
  PEOPLE_REL_TYPE,
  W15_NS,
} from "./constants";
import { withContentType } from "./contentTypes";
import type { CommentReferenceData, CommentReplyData } from "./model";

export interface ImportedPerson {
  author: string;
  providerId: string | null;
  userId: string | null;
  xml: string;
}

export interface ImportedPeople {
  partPath: string | null;
  xml: string | null;
  hadBom: boolean;
  ordered: readonly ImportedPerson[];
  /** The one person each name stands for: the one this editor recorded where several share a name */
  byAuthor: ReadonlyMap<string, ImportedPerson>;
}

export const NO_PEOPLE: ImportedPeople = {
  partPath: null,
  xml: null,
  hadBom: false,
  ordered: [],
  byAuthor: new Map(),
};

function attribute(el: Element, localName: string): string | null {
  return (
    Array.from(el.attributes).find((entry) => entry.localName === localName)
      ?.value ?? null
  );
}

function isOurs(person: ImportedPerson): boolean {
  return person.providerId === COMMENT_AUTHOR_PROVIDER;
}

/** Reads the people part related from the main document story. */
export function readPeople(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): ImportedPeople {
  const relationship = readRelationships(parts, relsPathOf(mainPartPath)).find(
    (entry) => entry.type === PEOPLE_REL_TYPE && !entry.external
  );
  if (!relationship) return NO_PEOPLE;
  const partPath = resolveTarget(mainPartPath, relationship.target);
  const bytes = parts.get(partPath);
  if (!bytes) return { ...NO_PEOPLE, partPath };

  const { text, hadBom } = decodeUtf8(bytes);
  const root = parseXml(text).documentElement;
  const ordered: ImportedPerson[] = [];
  const byAuthor = new Map<string, ImportedPerson>();
  for (const el of elementChildren(root)) {
    if (el.localName !== "person") continue;
    const author = attribute(el, "author");
    if (author === null) continue;
    const presence = elementChildren(el).find(
      (child) => child.localName === "presenceInfo"
    );
    const person: ImportedPerson = {
      author,
      providerId: presence ? attribute(presence, "providerId") : null,
      userId: presence ? attribute(presence, "userId") : null,
      xml: serializeXml(el),
    };
    ordered.push(person);
    const known = byAuthor.get(author);
    if (known === undefined || (!isOurs(known) && isOurs(person))) {
      byAuthor.set(author, person);
    }
  }
  return { partPath, xml: text, hadBom, ordered, byAuthor };
}

/** The identity this editor recorded for the author of that name. Null for a name it did not record */
export function commentAuthorId(
  people: ImportedPeople,
  author: string
): string | null {
  const person = people.byAuthor.get(author);
  return person !== undefined && isOurs(person) ? person.userId : null;
}

function personXml(author: string, userId: string): string {
  return (
    `<w15:person w15:author="${escapeXml(author)}">` +
    `<w15:presenceInfo w15:providerId="${COMMENT_AUTHOR_PROVIDER}" w15:userId="${escapeXml(userId)}"/>` +
    "</w15:person>"
  );
}

/** The authors whose identity the document has not recorded yet, each under the name they write as */
function unrecordedAuthors(
  bodies: Iterable<CommentReferenceData | CommentReplyData>,
  people: ImportedPeople
): Map<string, string> {
  const unrecorded = new Map<string, string>();
  for (const body of bodies) {
    if (body.author === null || body.authorId === null) continue;
    if (commentAuthorId(people, body.author) === body.authorId) continue;
    unrecorded.set(body.author, body.authorId);
  }
  return unrecorded;
}

function peopleXml(people: ImportedPeople, added: readonly string[]): string {
  if (people.xml === null) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w15:people xmlns:w15="${W15_NS}">${added.join("")}</w15:people>`
    );
  }
  const open = /<(?:[\w.-]+:)?people\b[^>]*>/.exec(people.xml);
  if (!open) {
    throw new DocxExportError(
      "malformed-xml",
      "the people part has no people root element"
    );
  }
  const name = /^<([^\s>]+)/.exec(open[0])?.[1];
  const close = name ? people.xml.lastIndexOf(`</${name}>`) : -1;
  if (close === -1) {
    throw new DocxExportError(
      "malformed-xml",
      "the people part has no closing people tag"
    );
  }
  return people.xml.slice(0, close) + added.join("") + people.xml.slice(close);
}

function availablePeoplePath(session: SessionStore): string {
  const directory = directoryOf(session.mainPartPath);
  for (let suffix = 0; ; suffix += 1) {
    const name = suffix === 0 ? "people.xml" : `people${suffix + 1}.xml`;
    const path = directory + name;
    if (!session.parts.has(path)) return path;
  }
}

/**
 * Plans the people part, its relationship and its content type for every identity the current
 * comments carry that the document has not recorded. Null when it has recorded them all, which
 * leaves the part as it arrived.
 */
export function planPeoplePart(
  bodies: Iterable<CommentReferenceData | CommentReplyData>,
  session: SessionStore,
  relationships: RelationshipWriter,
  currentContentTypes: Uint8Array | undefined
): ReadonlyMap<string, Uint8Array> | null {
  const people = session.comments.people;
  const added = Array.from(unrecordedAuthors(bodies, people), ([author, id]) =>
    personXml(author, id)
  );
  if (added.length === 0) return null;

  const parts = new Map<string, Uint8Array>();
  const addingPart = people.partPath === null;
  const partPath = people.partPath ?? availablePeoplePath(session);
  if (addingPart) {
    relationships.add({
      type: PEOPLE_REL_TYPE,
      target: partPath.slice(directoryOf(session.mainPartPath).length),
    });
  }
  parts.set(partPath, encodeUtf8(peopleXml(people, added), people.hadBom));
  if (addingPart || people.xml === null) {
    const contentTypes = withContentType(
      session.parts,
      partPath,
      PEOPLE_CONTENT_TYPE,
      currentContentTypes
    );
    if (contentTypes) parts.set(CONTENT_TYPES_PATH, contentTypes);
  }
  return parts;
}
