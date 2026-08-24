/**
 * Hyperlinks (`w:hyperlink`, §17.16.22): the wrapper a document brought in, and the relationship a
 * link's address lives on. `spec/notes/hyperlinks.md` has what the specification settles.
 *
 * A link names its target one of two ways. `r:id` points at a relationship whose target is the
 * address, so the address of an external link is never in the body at all; `w:anchor` names a
 * bookmark, which this editor does not model and therefore never reads.
 *
 * The wrapper travels as the verbatim opening tag, the way a content control's does (`docx/sdt`),
 * so `w:tooltip`, `w:history`, `w:docLocation` and anything else we never read go back out
 * untouched.
 */

import { attrString, R_NS } from "../ooxml/xml";
import {
  type RelationshipWriter,
  readRelationships,
  relsPathOf,
} from "./relationships";

const HYPERLINK_REL_TYPE = `${R_NS}/hyperlink`;

/** The one qualified name we read the relationship id under, and write it back as */
const REL_ID_ATTR = "r:id";

const REL_ID_VALUE = /\sr:id="([^"]*)"/;

export interface HyperlinkWrapper {
  /** The `<w:hyperlink>` opening tag as it came */
  prefix: string;
  /** The relationship its address lives on. null for a link that names a bookmark alone */
  relId: string | null;
}

/**
 * Takes a `w:hyperlink` apart into the wrapper to put back on export. null for a shape we do not
 * write ourselves, which leaves whatever held the link preserved as its own fragment.
 *
 * The `w:` prefix on the tag is one we write, and `r:id` is an attribute a retargeted link has
 * rewritten by name, so a document that binds either namespace to another prefix is left alone
 * rather than written back under ours.
 */
export function readHyperlinkWrapper(el: Element): HyperlinkWrapper | null {
  if (el.nodeName !== "w:hyperlink") return null;
  const relId = el.getAttributeNS(R_NS, "id");
  if (relId !== null && !el.hasAttribute(REL_ID_ATTR)) return null;
  const attrs = attrString(el);
  return {
    prefix: attrs ? `<w:hyperlink ${attrs}>` : "<w:hyperlink>",
    relId,
  };
}

/** The address of every external link the body can point at, keyed by relationship id */
export type LinkTargets = ReadonlyMap<string, string>;

export const NO_LINK_TARGETS: LinkTargets = new Map();

/**
 * Every address the body has a hyperlink relationship to.
 *
 * A relationship of another type, or one whose target is a part inside the package rather than an
 * address outside it, is left out: the link that points at it keeps its wrapper and offers no
 * address, the same as a link naming a bookmark.
 */
export function readLinkTargets(
  parts: Map<string, Uint8Array>,
  mainPartPath: string
): LinkTargets {
  const targets = new Map<string, string>();
  for (const rel of readRelationships(parts, relsPathOf(mainPartPath))) {
    if (rel.type !== HYPERLINK_REL_TYPE || !rel.external) continue;
    targets.set(rel.id, rel.target);
  }
  return targets;
}

/** The relationship id this opening tag names, null when it names none */
export function relIdIn(prefix: string): string | null {
  return REL_ID_VALUE.exec(prefix)?.[1] ?? null;
}

/**
 * The same opening tag pointing at this relationship. `<w:hyperlink>` on its own is what a link
 * made in the editor starts from.
 *
 * A tag that already names one has only that value rewritten, so a link nobody retargeted comes
 * out as the very string it went in as. A tag that named none is given the attribute, which the
 * specification has supersede any `w:anchor` beside it, along with the namespace it lives in: the
 * body we write into is not always one that declares it, and declaring it again where it is
 * already bound to the same namespace changes nothing.
 */
export function withRelId(prefix: string, relId: string): string {
  if (REL_ID_VALUE.test(prefix)) {
    return prefix.replace(REL_ID_VALUE, ` ${REL_ID_ATTR}="${relId}"`);
  }
  const name = "<w:hyperlink";
  return `${name} xmlns:r="${R_NS}" ${REL_ID_ATTR}="${relId}"${prefix.slice(name.length)}`;
}

/** Which relationship each address goes out on while the body is being written */
export interface LinkRefs {
  /**
   * The relationship the link at this address points at. `was` is the id the link came in on,
   * which is kept where it still leads to that same address.
   * undefined where there is no relationships part to add to, which is a serializer running
   * outside an export.
   */
  relIdOf(href: string, was: string | null): string | undefined;
}

/** What a serializer called on its own, outside an export, has to work with */
export const NO_LINK_REFS: LinkRefs = { relIdOf: () => undefined };

/**
 * Hands out the relationship for every address the body points at, adding one where the package
 * has none.
 *
 * Nothing is planned ahead: an address is asked for as its link is written, and the relationships
 * part is built once the body is (`docx/exportDocx`).
 * An address already in the package is pointed at again rather than written twice, and the
 * relationship a link arrived on is kept wherever it still leads where the link now goes, so
 * retargeting one link in a document leaves every other link's XML alone.
 */
export function hyperlinkRefs(relationships: RelationshipWriter): LinkRefs {
  const external = relationships.opened.filter(
    (rel) => rel.type === HYPERLINK_REL_TYPE && rel.external
  );
  const added = new Map<string, string>();
  return {
    relIdOf: (href, was) => {
      const kept = external.find((rel) => rel.id === was);
      if (kept?.target === href) return was ?? undefined;
      const found = external.find((rel) => rel.target === href);
      if (found) return found.id;
      const already = added.get(href);
      if (already !== undefined) return already;
      const id = relationships.add({
        type: HYPERLINK_REL_TYPE,
        target: href,
        external: true,
      });
      added.set(href, id);
      return id;
    },
  };
}
