/**
 * The people part, where Word records who a comment author is beyond the display name
 * (`w15:people`): a `w15:person` carrying a `w15:presenceInfo` whose `w15:providerId` names the
 * directory that issued `w15:userId`.
 *
 * This editor records the identity a host application hands it under a provider of its own
 * (`COMMENT_AUTHOR_PROVIDER`). The `w15:author` name is the key both Word and this editor read the
 * part by, so a name stands for one identity per file: a name the part already records is read as
 * it stands and never appended to. Appending a second person for it would leave a file naming two
 * identities for one name, where a reader keying by name hands one author's comments to the other.
 */

import { DocxExportError } from "../../ooxml/errors";
import {
  childByLocalName,
  decodeUtf8,
  elementChildren,
  encodeUtf8,
  escapeXml,
  parseXml,
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

export interface ImportedPeople {
  partPath: string | null;
  xml: string | null;
  hadBom: boolean;
  /**
   * Every author name the part records, mapped to the identity it stands for. Null where this
   * editor cannot vouch for the name: another provider recorded it, or it is recorded twice over
   * under different identities and nothing in the file says which of them wrote what.
   */
  byAuthor: ReadonlyMap<string, string | null>;
}

export const NO_PEOPLE: ImportedPeople = {
  partPath: null,
  xml: null,
  hadBom: false,
  byAuthor: new Map(),
};

function attribute(el: Element, localName: string): string | null {
  return (
    Array.from(el.attributes).find((entry) => entry.localName === localName)
      ?.value ?? null
  );
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
  const ourIds = new Map<string, Set<string | null>>();
  const byAuthor = new Map<string, string | null>();
  for (const el of elementChildren(root)) {
    if (el.localName !== "person") continue;
    const author = attribute(el, "author");
    if (author === null) continue;
    byAuthor.set(author, null);
    const presence = childByLocalName(el, "presenceInfo");
    if (presence === null) continue;
    if (attribute(presence, "providerId") !== COMMENT_AUTHOR_PROVIDER) continue;
    const ids = ourIds.get(author) ?? new Set<string | null>();
    ids.add(attribute(presence, "userId"));
    ourIds.set(author, ids);
  }
  for (const [author, ids] of ourIds) {
    if (ids.size !== 1) continue;
    byAuthor.set(author, Array.from(ids)[0] ?? null);
  }
  return { partPath, xml: text, hadBom, byAuthor };
}

/**
 * The identity this editor recorded for the author of that name. Null for a name it did not
 * record, and for one the part records under more than one identity.
 */
export function commentAuthorId(
  people: ImportedPeople,
  author: string
): string | null {
  return people.byAuthor.get(author) ?? null;
}

function personXml(
  author: string,
  userId: string,
  prefix: string,
  declaration: string
): string {
  return (
    `<${prefix}person${declaration} ${prefix}author="${escapeXml(author)}">` +
    `<${prefix}presenceInfo ${prefix}providerId="${COMMENT_AUTHOR_PROVIDER}" ${prefix}userId="${escapeXml(userId)}"/>` +
    `</${prefix}person>`
  );
}

/**
 * The authors the part records no person for at all, each under the name they write as. A name it
 * already records is left as it stands, whichever provider recorded it and whatever it resolves to.
 */
function unrecordedAuthors(
  bodies: Iterable<CommentReferenceData | CommentReplyData>,
  people: ImportedPeople
): Map<string, string> {
  const unrecorded = new Map<string, string>();
  for (const body of bodies) {
    if (body.author === null || body.authorId === null) continue;
    if (people.byAuthor.has(body.author)) continue;
    unrecorded.set(body.author, body.authorId);
  }
  return unrecorded;
}

/**
 * The prefix the root binds to the w15 namespace, ready to write in front of a name: empty when it
 * binds the namespace as the default one, null when the root binds it to nothing at all.
 */
function w15Prefix(root: Element): string | null {
  if (root.namespaceURI === W15_NS) {
    return root.prefix === null ? "" : `${root.prefix}:`;
  }
  const declaration = Array.from(root.attributes).find(
    (attr) =>
      attr.value === W15_NS &&
      (attr.name === "xmlns" || attr.name.startsWith("xmlns:"))
  );
  if (declaration === undefined) return null;
  const prefix = declaration.name.slice("xmlns:".length);
  return prefix === "" ? "" : `${prefix}:`;
}

/**
 * Where the root element's opening tag starts. The prolog is walked construct by construct so that
 * a `<` inside a comment or a processing instruction is not taken for the root. Every construct is
 * terminated, the part having been parsed already.
 */
function rootTagStart(xml: string): number {
  let at = 0;
  for (;;) {
    const opens = xml.indexOf("<", at);
    if (opens === -1) return -1;
    if (xml.startsWith("<!--", opens)) {
      at = xml.indexOf("-->", opens + 4) + 3;
      continue;
    }
    if (xml.startsWith("<?", opens)) {
      at = xml.indexOf("?>", opens + 2) + 2;
      continue;
    }
    return opens;
  }
}

/**
 * Where the opening tag that starts at `at` ends, and whether it closed the element on its own.
 * Quoted attribute values are stepped over, as a `>` inside one does not end the tag.
 */
function openTagEnd(
  xml: string,
  at: number
): { end: number; selfClosing: boolean } | null {
  let quote: string | null = null;
  for (let i = at; i < xml.length; i += 1) {
    const character = xml[i];
    if (quote !== null) {
      if (character === quote) quote = null;
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character === ">") {
      return { end: i + 1, selfClosing: xml[i - 1] === "/" };
    }
  }
  return null;
}

function malformed(detail: string): DocxExportError {
  return new DocxExportError("malformed-xml", `the people part ${detail}`);
}

/** The part with the new people spliced in, leaving everything it already held as it was. */
function peopleXml(
  people: ImportedPeople,
  added: ReadonlyMap<string, string>
): string {
  const xml = people.xml;
  if (xml === null) {
    const persons = Array.from(added, ([author, userId]) =>
      personXml(author, userId, "w15:", "")
    );
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w15:people xmlns:w15="${W15_NS}">${persons.join("")}</w15:people>`
    );
  }

  const root = parseXml(xml).documentElement;
  if (root.localName !== "people") {
    throw malformed("has no people root element");
  }
  const prefix = w15Prefix(root);
  const persons = Array.from(added, ([author, userId]) =>
    personXml(
      author,
      userId,
      prefix ?? "w15:",
      prefix === null ? ` xmlns:w15="${W15_NS}"` : ""
    )
  ).join("");

  const start = rootTagStart(xml);
  const open = start === -1 ? null : openTagEnd(xml, start);
  if (open === null) throw malformed("has no people root element");
  if (open.selfClosing) {
    return (
      `${xml.slice(0, open.end - 2)}>${persons}</${root.nodeName}>` +
      xml.slice(open.end)
    );
  }
  const close = xml.lastIndexOf(`</${root.nodeName}`);
  if (close === -1) throw malformed("has no closing people tag");
  return xml.slice(0, close) + persons + xml.slice(close);
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
  const added = unrecordedAuthors(bodies, people);
  if (added.size === 0) return null;

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
