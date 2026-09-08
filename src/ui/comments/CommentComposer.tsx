import {
  type FormEvent,
  type ReactElement,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
} from "react";
import type { CommentAuthor } from "../../editor/commands/commentCommands";
import { editorClassNames } from "../../styles/classNames";

interface FocusedTextareaProps {
  label: string;
  value: string;
  onChange: (value: string) => void;
}

export function FocusedTextarea({
  label,
  value,
  onChange,
}: FocusedTextareaProps): ReactElement {
  const input = useRef<HTMLTextAreaElement | null>(null);
  // Taking the focus must not take the page with it. A form standing in the comment rail is put
  // where its anchor is only once the rail has measured itself, so a browser scrolling to whatever
  // the focus landed on would scroll to where the form stood before it was placed - the top of the
  // document - and the reader would lose the text they were commenting on
  useLayoutEffect(() => input.current?.focus({ preventScroll: true }), []);
  useEffect(() => {
    const field = input.current;
    // jsdom draws nothing and has no `scrollIntoView`; the browser suite covers the scrolling
    if (typeof field?.scrollIntoView !== "function") return;
    const frame = requestAnimationFrame(() => {
      field.scrollIntoView({ block: "nearest" });
    });
    return () => cancelAnimationFrame(frame);
  }, []);
  return (
    <textarea
      ref={input}
      className={editorClassNames.commentInput}
      aria-label={label}
      value={value}
      onChange={(event) => onChange(event.target.value)}
    />
  );
}

export interface CommentComposerProps {
  author: CommentAuthor;
  label: string;
  submitLabel: string;
  /**
   * Writes the comment, and answers why it was not written where it was not. The text stays in
   * the form then, together with the reason, rather than the form standing there as if nothing
   * had been asked of it.
   */
  onSubmit: (text: string) => string | null;
  onClose: () => void;
}

export function CommentComposer({
  author,
  label,
  submitLabel,
  onSubmit,
  onClose,
}: CommentComposerProps): ReactElement {
  const [text, setText] = useState("");
  const [refused, setRefused] = useState<string | null>(null);
  const submit = (event: FormEvent) => {
    event.preventDefault();
    const reason = onSubmit(text);
    setRefused(reason);
    if (reason === null) onClose();
  };
  return (
    <form className={editorClassNames.commentComposer} onSubmit={submit}>
      <div className={editorClassNames.commentMeta}>{author.name}</div>
      <FocusedTextarea
        label={label}
        value={text}
        onChange={(next) => {
          setRefused(null);
          setText(next);
        }}
      />
      {refused !== null && (
        <p className={editorClassNames.commentMeta} role="alert">
          {refused}
        </p>
      )}
      <div className={editorClassNames.commentActions}>
        <button type="button" onClick={onClose}>
          Cancel
        </button>
        <button
          type="submit"
          className={editorClassNames.commentPrimaryAction}
          disabled={text.trim().length === 0}
        >
          {submitLabel}
        </button>
      </div>
    </form>
  );
}
