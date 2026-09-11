import { Fragment, type Node as PMNode } from "prosemirror-model";
import { spellNumber } from "../../numbering/spellers";
import { type StoryKey, storyKey } from "../../schema/stories";
import { type DocumentSection, sectionsOf } from "../sections";
import {
  type NoteKind,
  type NoteNumbering,
  type NoteNumberingProps,
  withNoteProps,
} from "./reading";

const NOTE_KINDS: readonly NoteKind[] = ["footnote", "endnote"];

/** A spelling longer than this is drawn as the count in decimal, which bounds a crafted `w:numStart` */
const MAX_LABEL_CHARS = 64;

interface Reference {
  readonly pos: number;
  readonly node: PMNode;
}

interface SectionReferences {
  readonly section: DocumentSection;
  readonly references: readonly Reference[];
}

/**
 * Every section of the body with the note references standing in its blocks, in document order,
 * and none at all for a body holding no reference.
 *
 * Reading the sections parses the properties of every paragraph, which a document without notes -
 * most of them - has no need to pay for on opening.
 */
function referencesBySection(doc: PMNode): readonly SectionReferences[] {
  const found: { readonly block: number; readonly reference: Reference }[] = [];
  doc.forEach((block, offset, index) => {
    block.descendants((node, pos) => {
      if (node.type.name === "noteReference") {
        found.push({
          block: index,
          reference: { pos: offset + 1 + pos, node },
        });
      }
      return true;
    });
  });
  if (found.length === 0) return [];
  const sections = sectionsOf(doc);
  const references: Reference[][] = sections.map(() => []);
  let at = 0;
  for (const { block, reference } of found) {
    while (block > sections[at].lastBlock) at += 1;
    references[at].push(reference);
  }
  return sections.map((section, index) => ({
    section,
    references: references[index],
  }));
}

function kindOf(node: PMNode): NoteKind {
  return node.attrs.kind === "endnote" ? "endnote" : "footnote";
}

function idOf(node: PMNode): string | null {
  const id: unknown = node.attrs.id;
  return typeof id === "string" ? id : null;
}

function propsIn(
  section: DocumentSection,
  numbering: NoteNumbering
): Readonly<Record<NoteKind, NoteNumberingProps>> {
  return {
    footnote: withNoteProps(numbering.footnote, section.props.footnotePr),
    endnote: withNoteProps(numbering.endnote, section.props.endnotePr),
  };
}

/** Where the count of one kind of note stands: the next number, and the label each id already took */
interface Count {
  next: number;
  readonly taken: Map<string, string>;
}

function labelOf(
  node: PMNode,
  count: Count,
  props: NoteNumberingProps,
  special: ReadonlySet<StoryKey>
): string {
  if (node.attrs.customMarkFollows === true) return "";
  const id = idOf(node);
  if (id === null || special.has(storyKey(kindOf(node), id))) return "?";
  const taken = count.taken.get(id);
  if (taken !== undefined) return taken;
  const label = spellNumber(count.next, props.format, MAX_LABEL_CHARS);
  count.next += 1;
  count.taken.set(id, label);
  return label;
}

/**
 * The label each note reference of the body is drawn with, by position.
 *
 * Each kind is counted on its own, in the order its references stand: a repeated id takes the
 * label of its first reference, a reference whose custom mark follows takes neither a label nor a
 * number (§17.11.14), and one calling a separator entry is no note to count. A section's own
 * properties lay over the settings, and a section that restarts each section counts from its start
 * again (§17.11.19, §17.11.20). A restart on each page is counted straight through, since the label
 * would then depend on the pages, which depend in turn on how wide the labels are drawn.
 *
 * It reads the document and what it is handed and nothing else, so the file an edited document
 * exports opens with the labels the editor drew.
 */
export function noteLabelsIn(
  doc: PMNode,
  numbering: NoteNumbering,
  special: ReadonlySet<StoryKey>
): ReadonlyMap<number, string> {
  const labels = new Map<number, string>();
  const counts: Record<NoteKind, Count> = {
    footnote: { next: 1, taken: new Map() },
    endnote: { next: 1, taken: new Map() },
  };
  referencesBySection(doc).forEach(({ section, references }, index) => {
    const props = propsIn(section, numbering);
    for (const kind of NOTE_KINDS) {
      if (index === 0 || props[kind].restart === "eachSect") {
        counts[kind].next = props[kind].start;
      }
    }
    for (const { pos, node } of references) {
      const kind = kindOf(node);
      labels.set(pos, labelOf(node, counts[kind], props[kind], special));
    }
  });
  return labels;
}

/**
 * What `noteLabelsIn` reads off the document, as one string: the kind, id, and custom mark of each
 * reference, which section it stands in, and what each section lays over the settings. Two
 * documents with the same signature are labelled alike under the same numbering.
 */
export function noteLabelSignature(doc: PMNode): string {
  return JSON.stringify(
    referencesBySection(doc).map(({ section, references }) => [
      section.props.footnotePr,
      section.props.endnotePr,
      references.map(({ node }) => [
        kindOf(node),
        idOf(node),
        node.attrs.customMarkFollows === true,
      ]),
    ])
  );
}

/**
 * The node with each label written onto the reference at its position, and the same node where
 * every label already stood.
 *
 * Built from nodes rather than steps, since `./core` reaches this and carries no transform package.
 */
function relabelled(
  node: PMNode,
  contentStart: number,
  labels: ReadonlyMap<number, string>
): PMNode {
  const children: PMNode[] = [];
  let changed = false;
  node.forEach((child, offset) => {
    const at = contentStart + offset;
    const label = labels.get(at);
    const next =
      label !== undefined && child.attrs.label !== label
        ? child.type.create({ ...child.attrs, label }, null, child.marks)
        : child.isLeaf
          ? child
          : relabelled(child, at + 1, labels);
    changed ||= next !== child;
    children.push(next);
  });
  return changed ? node.copy(Fragment.from(children)) : node;
}

/** The document with every note reference labelled, and the same node where every label already stood */
export function withNoteLabels(
  doc: PMNode,
  numbering: NoteNumbering,
  special: ReadonlySet<StoryKey>
): PMNode {
  const labels = noteLabelsIn(doc, numbering, special);
  return labels.size === 0 ? doc : relabelled(doc, 0, labels);
}
