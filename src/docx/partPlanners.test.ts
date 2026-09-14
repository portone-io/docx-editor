import { describe, expect, it } from "vitest";
import {
  PART_PLANNERS,
  STORY_ENTRIES_PARTS,
  STORY_WRITINGS,
} from "./partPlanners";
import { EVERY_STORY_CHANGE } from "./storyParts";

const plannerNames = () => PART_PLANNERS.map(({ name }) => name);

describe("the part planners", () => {
  it("runs one planner per writer and derives what is written from the same list", () => {
    expect(plannerNames()).toEqual([
      "numbering",
      "comments",
      "headers and footers",
      "footnotes",
    ]);
    expect(STORY_WRITINGS.map(({ kind, changes }) => [kind, changes])).toEqual([
      ["comment", EVERY_STORY_CHANGE],
      ["header", ["edited"]],
      ["footer", ["edited"]],
      ["footnote", EVERY_STORY_CHANGE],
    ]);
    expect(STORY_ENTRIES_PARTS.map(({ name }) => name)).toEqual(["footnotes"]);
  });

  it("names each story kind once, since a second writer of one kind would go unread", () => {
    const kinds = STORY_WRITINGS.map(({ kind }) => kind);

    expect(kinds).toEqual(Array.from(new Set(kinds)));
  });

  it("runs a planner for every part written one entry per story", () => {
    for (const part of STORY_ENTRIES_PARTS) {
      expect(plannerNames()).toContain(part.name);
      expect(
        STORY_WRITINGS.find(({ kind }) => kind === part.kind)?.changes
      ).toEqual(EVERY_STORY_CHANGE);
    }
  });
});
