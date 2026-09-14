// @vitest-environment jsdom
import { Transform } from "prosemirror-transform";
import { describe, expect, it } from "vitest";
import { makeNotesDocx } from "../__testing__/docx";
import { storyKey, storyNodeOf } from "../schema/stories";
import { importDocx } from "./importDocx";
import { setStory, storyFromText, withoutStories } from "./story";
import { type StoryChange, storyChangesOf } from "./storyParts";

/** What each change says, keyed the way a reader of the list would name the story */
function summary(changes: readonly StoryChange[]): readonly string[] {
  return changes.map((change) =>
    change.change === "added"
      ? `added ${change.key}`
      : `${change.change} ${change.imported.key}`
  );
}

describe("the changes to the stories of one kind", () => {
  it("reads a story nobody touched as kept", () => {
    const { doc, session } = importDocx(makeNotesDocx());

    expect(summary(storyChangesOf(doc, session, "footnote"))).toEqual([
      "kept footnote:-1",
      "kept footnote:2",
    ]);
  });

  it("reads a story whose source changed as edited and one the document dropped as removed", () => {
    const { doc, session } = importDocx(makeNotesDocx());
    const edited = withoutStories(
      setStory(
        new Transform(doc),
        storyKey("footnote", "2"),
        storyFromText("Rewritten")
      ),
      [storyKey("footnote", "-1")]
    ).doc;

    const changes = storyChangesOf(edited, session, "footnote");
    expect(summary(changes)).toEqual([
      "removed footnote:-1",
      "edited footnote:2",
    ]);
    const written = changes[1];
    expect(written?.change === "edited" && written.current).toBe(
      storyNodeOf(edited, storyKey("footnote", "2"))
    );
  });

  it("orders added stories after the part's own entries by id", () => {
    const { doc, session } = importDocx(makeNotesDocx());
    const tr = new Transform(doc);
    for (const id of ["10", "9", "-5"]) {
      setStory(tr, storyKey("footnote", id), storyFromText(`Note ${id}`));
    }

    expect(summary(storyChangesOf(tr.doc, session, "footnote"))).toEqual([
      "kept footnote:-1",
      "kept footnote:2",
      "added footnote:-5",
      "added footnote:9",
      "added footnote:10",
    ]);
    expect(summary(storyChangesOf(tr.doc, session, "endnote"))).toEqual([
      "kept endnote:0",
      "kept endnote:3",
    ]);
  });
});
