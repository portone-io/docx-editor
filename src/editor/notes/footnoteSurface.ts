/**
 * The note surface bound to footnotes, which are the notes an edit may add, rewrite and delete
 * (`schema/stories`).
 *
 * The surface itself is written for a kind of note, the way the note commands are
 * (`editor/commands/footnoteCommands`), so the endnotes it would take are one binding away. What
 * keeps the editor and the export from disagreeing is that nothing binds it to endnotes: the
 * Endnotes part goes back out as it arrived.
 */

import type { EditorView } from "prosemirror-view";
import { type StoryKey, storyKey } from "../../schema/stories";
import type { StoryExtensions, StoryHost } from "../stories/storyView";
import { noteExtensions, noteHost, noteIdIn } from "./noteSurface";

/** The id this story key names, and null for a key naming anything but a footnote */
export function footnoteIdOf(key: StoryKey): string | null {
  return noteIdIn("footnote", key);
}

/** Writes a footnote's edits into the main document and hands the caret back to its reference */
export function footnoteHost(
  main: EditorView,
  activate: StoryHost["activate"]
): StoryHost {
  return noteHost(main, "footnote", activate);
}

/** What one footnote's editing view draws its number as and what its Backspace deletes */
export function footnoteExtensions(
  main: EditorView,
  id: string,
  labelOf: () => string
): StoryExtensions {
  return noteExtensions(main, "footnote", storyKey("footnote", id), labelOf);
}
