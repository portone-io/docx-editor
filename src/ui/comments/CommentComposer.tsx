import {
  type FormEvent,
  type ReactElement,
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
  useLayoutEffect(() => input.current?.focus(), []);
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
  onSubmit: (text: string) => boolean;
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
  const submit = (event: FormEvent) => {
    event.preventDefault();
    if (onSubmit(text)) onClose();
  };
  return (
    <form className={editorClassNames.commentComposer} onSubmit={submit}>
      <div className={editorClassNames.commentMeta}>{author.name}</div>
      <FocusedTextarea label={label} value={text} onChange={setText} />
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
