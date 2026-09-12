// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { insertFootnote } from "../editor/commands/noteCommands";
import {
  EVERY_CAPABILITY,
  NO_CAPABILITY,
  type SurfaceCapabilities,
} from "../editor/stories/storyView";
import { footnoteItem } from "./footnoteItem";

/** A surface whose part carries a table and a link, and no note of its own */
const WITHOUT_NOTES: SurfaceCapabilities = new Set(["table", "link"]);

describe("the insert footnote row", () => {
  it("is there only where the surface says it takes a note", () => {
    expect(footnoteItem(EVERY_CAPABILITY)?.label).toBe("Insert footnote");
    // A note's own story takes nothing of the kind (`editor/notes/noteSurface`), and neither
    // does a surface that takes other content but no note
    expect(footnoteItem(NO_CAPABILITY)).toBeNull();
    expect(footnoteItem(WITHOUT_NOTES)).toBeNull();
  });

  it("names the one command both menus run and ask", () => {
    expect(footnoteItem(EVERY_CAPABILITY)?.command).toBe(insertFootnote);
  });
});
