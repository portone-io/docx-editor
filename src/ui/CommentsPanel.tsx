import { Check, Pencil, RotateCcw, Trash2 } from "lucide-react";
import type { Command, EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useMemo, useRef, useState } from "react";
import {
  addComment,
  addCommentReply,
  type CommentAuthor,
  type DocumentComment,
  documentComments,
  removeComment,
  removeCommentReply,
  selectComment,
  setCommentResolved,
  updateComment,
  updateCommentReply,
} from "../editor/commands/commentCommands";
import { editorClassNames } from "../styles/classNames";
import { CommentComposer, FocusedTextarea } from "./comments/CommentComposer";
import {
  COMPOSER_POSITION,
  useCommentRailLayout,
} from "./comments/useCommentRailLayout";
import { tooltipAttribute } from "./Tooltip";

function shownDate(value: string | null): string | null {
  if (value === null) return null;
  const date = new Date(value);
  if (Number.isNaN(date.valueOf())) return value;
  return `${date.toISOString().slice(0, 16).replace("T", " ")} UTC`;
}

function run(view: EditorView, command: Command): boolean {
  const applied = command(
    view.state,
    (transaction) => view.dispatch(transaction),
    view
  );
  if (applied) view.focus();
  return applied;
}

interface EditTarget {
  commentId: string;
  replyId: string | null;
}

function sameTarget(left: EditTarget | null, right: EditTarget): boolean {
  return left?.commentId === right.commentId && left.replyId === right.replyId;
}

export interface CommentsPanelProps {
  view: EditorView;
  state: EditorState;
  readOnly: boolean;
  composerOpen: boolean;
  author: CommentAuthor;
  closeComposer: () => void;
  scrollContainer: HTMLElement | null;
  allCommentsOpen: boolean;
}

export function CommentsPanel({
  view,
  state,
  readOnly,
  composerOpen,
  author,
  closeComposer,
  scrollContainer,
  allCommentsOpen,
}: CommentsPanelProps) {
  // biome-ignore lint/correctness/useExhaustiveDependencies: comment data changes with the document, not with selection-only transactions
  const comments = useMemo(() => documentComments(state), [state.doc]);
  const panel = useRef<HTMLElement | null>(null);
  const [editing, setEditing] = useState<EditTarget | null>(null);
  const [draft, setDraft] = useState("");
  const [replying, setReplying] = useState<string | null>(null);
  const composerSelectionPos = composerOpen ? state.selection.from : -1;
  const visible = useMemo(
    () =>
      allCommentsOpen
        ? [...comments].sort((left, right) => {
            const leftTime = Date.parse(left.date ?? "");
            const rightTime = Date.parse(right.date ?? "");
            if (Number.isNaN(leftTime) && Number.isNaN(rightTime)) {
              return right.referencePos - left.referencePos;
            }
            if (Number.isNaN(leftTime)) return 1;
            if (Number.isNaN(rightTime)) return -1;
            return rightTime - leftTime;
          })
        : comments.filter((comment) => !comment.resolved),
    [allCommentsOpen, comments]
  );

  const { positions, canvasHeight } = useCommentRailLayout({
    panel,
    view,
    scrollContainer,
    visible,
    composerOpen,
    composerSelectionPos,
    allCommentsOpen,
  });

  const beginEdit = (
    commentId: string,
    replyId: string | null,
    text: string
  ) => {
    setReplying(null);
    setEditing({ commentId, replyId });
    setDraft(text);
  };

  const saveEdit = (comment: DocumentComment, replyId: string | null) => {
    const original =
      replyId === null
        ? comment.text
        : (comment.replies.find((reply) => reply.id === replyId)?.text ?? "");
    const saved =
      draft === original ||
      run(
        view,
        replyId === null
          ? updateComment(comment.id, draft)
          : updateCommentReply(comment.id, replyId, draft)
      );
    if (saved) {
      setEditing(null);
      setDraft("");
    }
  };

  return (
    <aside
      ref={panel}
      className={editorClassNames.commentsPanel}
      aria-label="Comments"
      data-view={allCommentsOpen ? "all" : "rail"}
      style={allCommentsOpen ? undefined : { height: `${canvasHeight}px` }}
    >
      {allCommentsOpen && (
        <div className={editorClassNames.commentsHeading}>
          <h2>All comments</h2>
        </div>
      )}
      <div
        className={editorClassNames.commentsCanvas}
        style={{ height: allCommentsOpen ? undefined : `${canvasHeight}px` }}
      >
        {composerOpen && !readOnly && (
          <div
            className={editorClassNames.commentPosition}
            data-comment-position={COMPOSER_POSITION}
            style={
              allCommentsOpen
                ? undefined
                : {
                    top: `${positions.get(COMPOSER_POSITION) ?? 64}px`,
                  }
            }
          >
            <CommentComposer
              author={author}
              label="Comment text"
              submitLabel="Comment"
              onClose={closeComposer}
              onSubmit={(text) =>
                run(
                  view,
                  addComment({
                    text,
                    author: author.name,
                    initials: author.initials,
                  })
                )
              }
            />
          </div>
        )}
        {visible.map((comment) => {
          const rootTarget = { commentId: comment.id, replyId: null };
          const rootEditing = !readOnly && sameTarget(editing, rootTarget);
          return (
            <article
              className={editorClassNames.commentPosition}
              data-resolved={comment.resolved ? "true" : undefined}
              data-comment-position={comment.id}
              key={comment.id}
              style={
                allCommentsOpen
                  ? undefined
                  : { top: `${positions.get(comment.id) ?? 64}px` }
              }
            >
              <div className={editorClassNames.commentCard}>
                <div className={editorClassNames.commentHeader}>
                  <button
                    type="button"
                    className={editorClassNames.commentMeta}
                    onClick={() => run(view, selectComment(comment.id))}
                    disabled={comment.resolved}
                  >
                    <span className={editorClassNames.commentAuthor}>
                      {comment.author ?? "Unknown author"}
                    </span>
                    {shownDate(comment.date) && (
                      <span className={editorClassNames.commentDate}>
                        {shownDate(comment.date)}
                      </span>
                    )}
                  </button>
                  {!readOnly && !rootEditing && (
                    <div className={editorClassNames.commentIconActions}>
                      {!comment.resolved && (
                        <button
                          type="button"
                          aria-label="Edit"
                          {...tooltipAttribute("Edit")}
                          onClick={() =>
                            beginEdit(comment.id, null, comment.text)
                          }
                        >
                          <Pencil size={16} aria-hidden="true" />
                        </button>
                      )}
                      <button
                        type="button"
                        aria-label={comment.resolved ? "Reopen" : "Resolve"}
                        {...tooltipAttribute(
                          comment.resolved ? "Reopen" : "Resolve"
                        )}
                        onClick={() => {
                          setReplying(null);
                          setEditing(null);
                          run(
                            view,
                            setCommentResolved(comment.id, !comment.resolved)
                          );
                        }}
                      >
                        {comment.resolved ? (
                          <RotateCcw size={16} aria-hidden="true" />
                        ) : (
                          <Check size={16} aria-hidden="true" />
                        )}
                      </button>
                      {!comment.resolved && (
                        <button
                          type="button"
                          aria-label="Delete"
                          {...tooltipAttribute("Delete")}
                          onClick={() => run(view, removeComment(comment.id))}
                        >
                          <Trash2 size={16} aria-hidden="true" />
                        </button>
                      )}
                    </div>
                  )}
                </div>
                {rootEditing ? (
                  <FocusedTextarea
                    label="Edit comment"
                    value={draft}
                    onChange={setDraft}
                  />
                ) : readOnly || comment.resolved ? (
                  <p className={editorClassNames.commentBody}>{comment.text}</p>
                ) : (
                  <button
                    type="button"
                    className={editorClassNames.commentBody}
                    aria-label={`Reply to comment: ${comment.text}`}
                    onClick={() => {
                      setEditing(null);
                      setReplying(comment.id);
                    }}
                  >
                    {comment.text}
                  </button>
                )}
                {!readOnly && (
                  <div className={editorClassNames.commentActions}>
                    {rootEditing ? (
                      <>
                        <button type="button" onClick={() => setEditing(null)}>
                          Cancel
                        </button>
                        <button
                          type="button"
                          className={editorClassNames.commentPrimaryAction}
                          disabled={draft.trim().length === 0}
                          onClick={() => saveEdit(comment, null)}
                        >
                          Save
                        </button>
                      </>
                    ) : null}
                  </div>
                )}
                {comment.replies.map((reply) => {
                  const target = {
                    commentId: comment.id,
                    replyId: reply.id,
                  };
                  const isEditing = !readOnly && sameTarget(editing, target);
                  return (
                    <div
                      className={editorClassNames.commentReply}
                      key={reply.id}
                    >
                      <div className={editorClassNames.commentMeta}>
                        <span className={editorClassNames.commentAuthor}>
                          {reply.author ?? "Unknown author"}
                        </span>
                        {shownDate(reply.date) && (
                          <span className={editorClassNames.commentDate}>
                            {shownDate(reply.date)}
                          </span>
                        )}
                      </div>
                      {isEditing ? (
                        <FocusedTextarea
                          label="Edit reply"
                          value={draft}
                          onChange={setDraft}
                        />
                      ) : (
                        <p className={editorClassNames.commentBody}>
                          {reply.text}
                        </p>
                      )}
                      {!readOnly && !comment.resolved && (
                        <div className={editorClassNames.commentActions}>
                          {isEditing ? (
                            <>
                              <button
                                type="button"
                                onClick={() => setEditing(null)}
                              >
                                Cancel
                              </button>
                              <button
                                type="button"
                                className={
                                  editorClassNames.commentPrimaryAction
                                }
                                disabled={draft.trim().length === 0}
                                onClick={() => saveEdit(comment, reply.id)}
                              >
                                Save
                              </button>
                            </>
                          ) : (
                            <div
                              className={editorClassNames.commentIconActions}
                            >
                              <button
                                type="button"
                                aria-label="Edit reply"
                                {...tooltipAttribute("Edit reply")}
                                onClick={() =>
                                  beginEdit(comment.id, reply.id, reply.text)
                                }
                              >
                                <Pencil size={16} aria-hidden="true" />
                              </button>
                              <button
                                type="button"
                                aria-label="Delete reply"
                                {...tooltipAttribute("Delete reply")}
                                onClick={() =>
                                  run(
                                    view,
                                    removeCommentReply(comment.id, reply.id)
                                  )
                                }
                              >
                                <Trash2 size={16} aria-hidden="true" />
                              </button>
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}
                {replying === comment.id && !readOnly && !comment.resolved && (
                  <CommentComposer
                    author={author}
                    label="Reply text"
                    submitLabel="Reply"
                    onClose={() => setReplying(null)}
                    onSubmit={(text) =>
                      run(
                        view,
                        addCommentReply(comment.id, {
                          text,
                          author: author.name,
                          initials: author.initials,
                        })
                      )
                    }
                  />
                )}
              </div>
            </article>
          );
        })}
        {allCommentsOpen && visible.length === 0 && !composerOpen && (
          <p className={editorClassNames.commentsEmpty}>
            This document has no comments.
          </p>
        )}
      </div>
    </aside>
  );
}
