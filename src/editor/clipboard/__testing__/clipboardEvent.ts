/**
 * The paste event ProseMirror builds when a caller runs the paste logic itself
 * (`view.pasteHTML`, `view.pasteText`) and hands over no event of its own.
 *
 * jsdom defines no `ClipboardEvent`, so without this the call raises before any handler runs.
 * The event carries no clipboard, which is what a real one built the same way carries: every
 * handler that reaches for it has to hold that case anyway.
 */
class EmptyClipboardEvent extends Event implements ClipboardEvent {
  readonly clipboardData: DataTransfer | null = null;
}

/** Defines `ClipboardEvent` for a test environment that has none. Does nothing where one exists */
export function defineClipboardEvent(): void {
  if ("ClipboardEvent" in globalThis) return;
  globalThis.ClipboardEvent = EmptyClipboardEvent;
}
