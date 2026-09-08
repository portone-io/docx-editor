/**
 * The paste event ProseMirror builds when a caller runs the paste logic itself
 * (`view.pasteHTML`, `view.pasteText`) and hands over no event of its own.
 *
 * jsdom defines no `ClipboardEvent`, so without this the call raises before any handler runs.
 * Like the browser's constructor, it carries no clipboard unless one is supplied. The menu uses
 * a DataTransfer to pass both formats through the same handlers as a native paste.
 */
import { vi } from "vitest";

class TestClipboardEvent extends Event implements ClipboardEvent {
  readonly clipboardData: DataTransfer | null;

  constructor(type: string, options: ClipboardEventInit = {}) {
    super(type, options);
    this.clipboardData = options.clipboardData ?? null;
  }
}

/** The text-only transfer the menu builds. Tests carrying files supply their own event data */
class TextDataTransfer {
  private readonly data = new Map<string, string>();
  readonly files: File[] = [];

  get types(): string[] {
    return [...this.data.keys()];
  }

  setData(type: string, value: string): void {
    this.data.set(type, value);
  }

  getData(type: string): string {
    return this.data.get(type) ?? "";
  }
}

/** Defines `ClipboardEvent` for a test environment that has none. Does nothing where one exists */
export function defineClipboardEvent(): void {
  if (!("ClipboardEvent" in globalThis)) {
    globalThis.ClipboardEvent = TestClipboardEvent;
  }
  if (!("DataTransfer" in globalThis))
    vi.stubGlobal("DataTransfer", TextDataTransfer);
}
