// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { insertEndnote, insertFootnote } from "../editor/commands/noteCommands";
import { NOTE_KEYS } from "../editor/plugins/keymap";
import {
  EVERY_CAPABILITY,
  NO_CAPABILITY,
  type SurfaceCapabilities,
} from "../editor/stories/storyView";
import { noteItems } from "./noteItems";

/** A surface whose part carries a table and a link, and no note of its own */
const WITHOUT_NOTES: SurfaceCapabilities = new Set(["table", "link"]);

describe("the rows that put a note in", () => {
  it("offers a row per kind only where the surface says it takes a note", () => {
    expect(noteItems(EVERY_CAPABILITY).map(({ label }) => label)).toEqual([
      "Insert footnote",
      "Insert endnote",
    ]);
    // A note's own story takes nothing of the kind (`editor/notes/noteSurface`), and neither
    // does a surface that takes other content but no note
    expect(noteItems(NO_CAPABILITY)).toEqual([]);
    expect(noteItems(WITHOUT_NOTES)).toEqual([]);
  });

  it("names the commands both menus run and ask", () => {
    expect(noteItems(EVERY_CAPABILITY).map(({ command }) => command)).toEqual([
      insertFootnote,
      insertEndnote,
    ]);
  });

  /**
   * The endnote is not on one key everywhere (`editor/plugins/keymap`), so a row that spelled its
   * own letter would promise a key the editor does not hold on one of the platforms.
   */
  it("labels each row with the letter its own binding ends on", () => {
    const hints = noteItems(EVERY_CAPABILITY).map(({ hint }) => hint);

    expect(hints[0]?.endsWith("F")).toBe(true);
    expect(NOTE_KEYS.footnote.endsWith("f")).toBe(true);
    expect(hints[1]?.endsWith(NOTE_KEYS.endnote.slice(-1).toUpperCase())).toBe(
      true
    );
  });
});
