/**
 * What a protection level lets a change rewrite, as one value every side of the package reads.
 *
 * A level says three things: which parts of the package a change may write, what an entry in one
 * of those parts may look like, and how the story reads once the markup the level is about is
 * taken out of it. The part planners write those parts and the verifier judges them, and each
 * carried its own copy of the first two with nothing but a test holding them together. Declaring
 * them once means a part the writer starts adding is a part the verifier already knows.
 */

import type { EditableComments } from "../schema/protection";
import type { SessionStore } from "./session";
import type { Story } from "./storyProjection";

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
