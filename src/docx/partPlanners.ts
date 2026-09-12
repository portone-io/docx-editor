/**
 * The parts written beside the body, and what each of their planners carries into the file.
 *
 * `docx/exportDocx` runs these planners and `docx/invariants` reads what they write off the same
 * list, so the two cannot disagree about it: a kind of story whose planner is missing, or a change
 * its planner leaves out, is refused before the export rather than dropped without a word.
 */

import { HEADER_FOOTER_KINDS } from "../schema/stories";
import { commentsPlanner } from "./comments";
import { headerFooterPlanner } from "./headersFooters";
import { FOOTNOTES_PART, footnotesPlanner } from "./notes/writing";
import { numberingPlanner } from "./numberingPlanner";
import type { PartPlanner } from "./partPlan";
import {
  EVERY_STORY_CHANGE,
  type StoryEntriesPart,
  type StoryWriting,
  storyEntriesWriting,
  storyWriting,
} from "./storyParts";

/** One planner, beside the side stories it writes and the part it writes them one entry apiece into */
interface StoryWriter {
  readonly planner: PartPlanner;
  readonly writes: readonly StoryWriting[];
  /** The part holding one entry per story, and null for a planner that writes no such part */
  readonly entriesPart: StoryEntriesPart | null;
}

function entriesWriter(
  planner: PartPlanner,
  part: StoryEntriesPart
): StoryWriter {
  return { planner, writes: [storyEntriesWriting(part)], entriesPart: part };
}

/**
 * In the order the parts go into the package.
 *
 * A header or footer part holds one story, so the writer rewrites the part of a story an edit
 * changed and has nowhere to put one added or removed: no relationship and no section reference
 * would name a part it created, and none would stop naming one it dropped.
 */
const PART_WRITERS: readonly StoryWriter[] = [
  { planner: numberingPlanner, writes: [], entriesPart: null },
  {
    planner: commentsPlanner,
    writes: [storyWriting("comment", EVERY_STORY_CHANGE)],
    entriesPart: null,
  },
  {
    planner: headerFooterPlanner,
    writes: HEADER_FOOTER_KINDS.map((kind) => storyWriting(kind, ["edited"])),
    entriesPart: null,
  },
  entriesWriter(footnotesPlanner, FOOTNOTES_PART),
];

export const PART_PLANNERS: readonly PartPlanner[] = PART_WRITERS.map(
  ({ planner }) => planner
);

/** What every planner writes back of the side stories the document holds */
export const STORY_WRITINGS: readonly StoryWriting[] = PART_WRITERS.flatMap(
  ({ writes }) => writes
);

/** The parts written one entry per story, which the identity pass walks as one part apiece */
export const STORY_ENTRIES_PARTS: readonly StoryEntriesPart[] =
  PART_WRITERS.flatMap(({ entriesPart }) =>
    entriesPart === null ? [] : [entriesPart]
  );
