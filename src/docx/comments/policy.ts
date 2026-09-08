/**
 * The protection that takes comments and nothing else, as the policy both the part planners and
 * the server verifier read (`../protectionPolicy`).
 */

import {
  commentAdditionsBy,
  commentEditsOwned,
  commentIdentitiesKept,
  unattributedCommentAuthors,
  withoutComments,
} from "../../schema/protection";
import {
  type ChangeVerdict,
  type PolicyOptions,
  registerProtectionPolicy,
} from "../protectionPolicy";
import { comparableStory, type Story } from "../storyProjection";
import { COMMENT_STORY_PARTS } from "./parts";

type StoryReason =
  | "body-changed"
  | "comment-not-owned"
  | "comment-author-forged";

type PartReason = "comment-markup-rejected";

const refused = (
  reason: StoryReason
): ChangeVerdict<StoryReason, PartReason> => ({ ok: false, reason });

/**
 * Whether the two stories say the same thing once the comments are taken out of them.
 *
 * The two files were written by different hands, so they are compared as this editor's writer
 * puts them out rather than as they are worded (`../storyProjection`). A story the writer cannot
 * put out at all is answered the way a changed one is: there is nothing to compare it against.
 * No file this package opens reaches that answer today, since every attr the writer needs is set
 * on import; it is here so that a story it cannot write is refused rather than thrown over.
 */
function sameBody(before: Story, after: Story): boolean {
  const was = comparableStory(before, withoutComments);
  const now = comparableStory(after, withoutComments);
  return (
    was !== null &&
    now !== null &&
    was.length === now.length &&
    was.every((block, at) => block === now[at])
  );
}

function commentsStoryKept(
  before: Story,
  after: Story,
  authorId: string,
  { editableComments }: PolicyOptions
): ChangeVerdict<StoryReason, PartReason> {
  if (!sameBody(before, after)) return refused("body-changed");
  if (
    !commentIdentitiesKept(before.doc, after.doc) ||
    !commentAdditionsBy(
      before.doc,
      after.doc,
      authorId,
      unattributedCommentAuthors(before.session.comments.ordered)
    )
  ) {
    return refused("comment-author-forged");
  }
  if (
    !commentEditsOwned(before.doc, after.doc, {
      protection: "comments",
      authorId,
      editableComments,
    })
  ) {
    return refused("comment-not-owned");
  }
  return { ok: true };
}

export const commentsPolicy = registerProtectionPolicy<StoryReason, PartReason>(
  {
    level: "comments",
    parts: COMMENT_STORY_PARTS,
    rejectedMarkup: "comment-markup-rejected",
    storyKept: commentsStoryKept,
  }
);
