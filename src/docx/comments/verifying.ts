/**
 * Whether every entry of the three comment parts is one this editor wrote, for an author who was
 * in a position to write it.
 *
 * The parts a comment is written across are the ones a submission may rewrite, so the package
 * comparison (`../commentOnlyChange`) leaves their bytes alone. Something has to read them, or a
 * file could carry anything at all inside a comment and still be answered as unchanged. What each
 * of those parts holds and what an entry in it is held to is declared beside them (`./parts`).
 */

import type { EditableComments } from "../../schema/protection";
import type { CommentOnlyVerdict } from "../commentOnlyChange";
import type { StoryPartKind } from "../protectionPolicy";
import type { Story } from "../storyProjection";
import { COMMENT_STORY_PARTS } from "./parts";

/**
 * Whether every entry of one part came back as it was, or as one this editor writes for an author
 * who could have written it.
 *
 * An entry the file no longer stands behind is one no edit through the editor could have reached,
 * and an entry it did not stand behind when it left is one no edit could have taken away. Both
 * halves hold for each of the three parts: a comment nothing refers to, thread state for no
 * comment, an identity for a name nobody writes under.
 */
function partKept(
  kind: StoryPartKind,
  before: Story,
  after: Story,
  authorId: string,
  editableComments: EditableComments
): boolean {
  const arrived = kind.entriesIn(before.session, "arrived") ?? new Map();
  const submitted = kind.entriesIn(after.session, "submitted");
  if (submitted === null) return false;

  const stoodBehindNow = kind.referents(after);
  for (const [id, entry] of submitted) {
    const original = arrived.get(id);
    if (original && original.xml === entry.xml) continue;
    if (!stoodBehindNow.has(id)) return false;
    if (original && kind.anyonesChange(entry.el, original.el)) continue;
    if (
      !kind.wellFormed(entry.el) ||
      !kind.allowed(
        entry.el,
        original?.el ?? null,
        authorId,
        { editableComments },
        after.session
      )
    ) {
      return false;
    }
  }

  const stoodBehindBefore = kind.referents(before);
  return Array.from(arrived.keys()).every(
    (id) => stoodBehindBefore.has(id) || submitted.has(id)
  );
}

/**
 * Whether the three comment parts came back holding only entries this editor writes, each of them
 * one this author was in a position to write.
 *
 * This runs after the story has been judged, so a comment rewritten, re-anchored or deleted by the
 * wrong hand is already named for what it is (`../commentOnlyChange`). What is left to this is
 * everything the story cannot show: markup forged into a body, an entry nothing refers to, an
 * identity recorded for somebody else.
 */
export function commentPartsKept(
  before: Story,
  after: Story,
  authorId: string,
  editableComments: EditableComments
): CommentOnlyVerdict {
  for (const kind of COMMENT_STORY_PARTS) {
    const path = kind.pathIn(after.session) ?? kind.pathIn(before.session);
    if (path === null) continue;
    if (!partKept(kind, before, after, authorId, editableComments)) {
      return { ok: false, reason: "comment-markup-rejected", part: path };
    }
  }
  return { ok: true };
}
