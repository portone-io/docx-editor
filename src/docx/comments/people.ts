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
 * A name the file's comments already write under carries the same weight, recorded or not: a
 * person appended for it would claim those comments too.
 */

import { xmlnsAttr } from "../../ooxml/element";
import { xmlnsDecl } from "../../ooxml/names";
import { splicePart } from "../../ooxml/partSplice";
import {
  attributeByLocalName,
  childByLocalName,
  decodeUtf8,
  elementChildren,
  encodeUtf8,
  parseXml,
} from "../../ooxml/xml";
import { relatedPartPath } from "../packageParts";
import type { PartPlanContext } from "../partPlan";
import type { StoryPartKind } from "../protectionPolicy";
import { directoryOf } from "../relationships";
import type { SessionStore } from "../session";
import { COMMENT_AUTHOR_PROVIDER, PEOPLE_REL_TYPE, W15_NS } from "./constants";
import { renderPerson } from "./grammar";
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

/** Reads the people part related from the main document story. */
export function readPeople(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): ImportedPeople {
  const partPath = relatedPartPath(parts, mainPartPath, PEOPLE_REL_TYPE);
  if (partPath === null) return NO_PEOPLE;
  const bytes = parts.get(partPath);
  if (!bytes) return { ...NO_PEOPLE, partPath };

  const { text, hadBom } = decodeUtf8(bytes);
  const root = parseXml(text).documentElement;
  const ourIds = new Map<string, Set<string | null>>();
  const byAuthor = new Map<string, string | null>();
  for (const el of elementChildren(root)) {
    if (el.localName !== "person") continue;
    const author = attributeByLocalName(el, "author");
    if (author === null) continue;
    byAuthor.set(author, null);
    const presence = childByLocalName(el, "presenceInfo");
    if (presence === null) continue;
    if (
      attributeByLocalName(presence, "providerId") !== COMMENT_AUTHOR_PROVIDER
    )
      continue;
    const ids = ourIds.get(author) ?? new Set<string | null>();
    ids.add(attributeByLocalName(presence, "userId"));
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

/**
 * The authors to record a person for, each under the name they write as: the ones the comments
 * name one identity for and the part does not record already. A name the part records is left as
 * it stands, whichever provider recorded it and whatever it resolves to. So is a name the comments
 * carry under more than one identity, one identity and none among them: recording it would hand
 * the comments carrying no identity to whoever the name then resolved to, and a document saved by
 * a Word that wrote no people part carries every one of its comments that way.
 */
export function unrecordedAuthors(
  bodies: Iterable<CommentReferenceData | CommentReplyData>,
  people: ImportedPeople
): Map<string, string> {
  const identities = new Map<string, Set<string | null>>();
  for (const body of bodies) {
    if (body.author === null) continue;
    if (people.byAuthor.has(body.author)) continue;
    const ids = identities.get(body.author) ?? new Set<string | null>();
    ids.add(body.authorId);
    identities.set(body.author, ids);
  }

  const unrecorded = new Map<string, string>();
  for (const [author, ids] of identities) {
    if (ids.size !== 1) continue;
    const only = Array.from(ids)[0] ?? null;
    if (only !== null) unrecorded.set(author, only);
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

/** The part with the new people spliced in, leaving everything it already held as it was. */
function peopleXml(
  people: ImportedPeople,
  added: ReadonlyMap<string, string>
): string {
  const xml = people.xml;
  if (xml === null) {
    const persons = Array.from(added, ([author, userId]) =>
      renderPerson(author, userId, "w15:", null)
    );
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w15:people ${xmlnsDecl("w15")}>${persons.join("")}</w15:people>`
    );
  }

  const prefix = w15Prefix(parseXml(xml).documentElement);
  const persons = Array.from(added, ([author, userId]) =>
    renderPerson(
      author,
      userId,
      prefix ?? "w15:",
      prefix === null ? xmlnsAttr("w15") : null
    )
  ).join("");
  return splicePart(xml, { root: "people", append: persons });
}

/**
 * Plans the people part, its relationship and its content type for every identity the current
 * comments carry that the document has not recorded. Null when it has recorded them all, which
 * leaves the part as it arrived.
 *
 * `part` is the description the verifier reads the part back through (`./parts`), so where this
 * writes it and what it declares it as are the same facts on both sides.
 */
export function planPeoplePart(
  part: StoryPartKind,
  bodies: Iterable<CommentReferenceData | CommentReplyData>,
  session: SessionStore,
  context: PartPlanContext
): ReadonlyMap<string, Uint8Array> | null {
  const people = session.comments.people;
  const added = unrecordedAuthors(bodies, people);
  if (added.size === 0) return null;

  const parts = new Map<string, Uint8Array>();
  const addingPart = part.pathIn(session) === null;
  const partPath = part.writePathIn(session);
  if (addingPart) {
    context.relationships.add({
      type: part.relType,
      target: partPath.slice(directoryOf(session.mainPartPath).length),
    });
  }
  parts.set(partPath, encodeUtf8(peopleXml(people, added), people.hadBom));
  if (addingPart || people.xml === null) {
    context.contentTypes.addOverride(partPath, part.contentType);
  }
  return parts;
}
