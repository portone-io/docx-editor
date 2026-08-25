/**
 * The surface the harness page hands the Playwright tests.
 *
 * It exists so that a test can read the document model and the drawn sheet from the same page,
 * and it is never part of what the package ships.
 */

export interface CompositionCounts {
  start: number;
  update: number;
  end: number;
}

/** One body block, both as the document holds it and as the browser draws it */
export interface BlockReport {
  index: number;
  pos: number;
  type: string;
  /** The text the document model holds */
  docText: string;
  /** The text on screen, which during a composition also carries the composing text */
  domText: string;
  /** The list marker drawn on the block, which no page break may take away */
  marker: string | null;
}

/** Where the caret is drawn, in the coordinates the page marks are drawn in */
export interface CaretBox {
  top: number;
  bottom: number;
  left: number;
}

export interface SelectionReport {
  from: number;
  to: number;
  anchor: number;
  head: number;
}

export interface DocxHarness {
  blocks(): BlockReport[];
  /** The whole document as text, one line per block */
  text(): string;
  /** Whether an IME composition is open right now */
  composing(): boolean;
  compositions(): CompositionCounts;
  /** The page pushes the layout has decorated the sheet with, as one comparable string */
  pushes(): string;
  /** All layout marks that move content between page bodies */
  pagination(): string;
  /** The spaces the layout has opened at the page breaks, in the same shape as `pushes` */
  spaces(): string;
  /** Where the caret stands, so a test can tell which page it is on */
  caretBox(): CaretBox;
  selection(): SelectionReport;
  /**
   * The height a block is drawn at. A block measuring nothing is left out of the sheet
   * measurement altogether (`src/page/measureBlocks`), so it is what decides whether a break
   * inside it is given a space
   */
  blockHeight(blockIndex: number): number;
  /** Puts the caret `offset` characters into a block's content, and answers the position used */
  caretAt(blockIndex: number, offset: number): number;
  /**
   * Selects `length` characters `offset` characters into one top-level block, and answers the
   * position the stretch starts at
   */
  selectText(blockIndex: number, offset: number, length: number): number;
  /** Puts the caret in the first cell of the first table, and answers the text that cell holds */
  caretInCell(): string;
  /**
   * Asks for the right click menu over whatever the caret is in.
   *
   * The menu is opened by a `contextmenu` event, which is what the browser sends for the context
   * menu key as well as for the right button. CDP will not synthesize that key
   * (`Input.dispatchKeyEvent` produces no such event), so the event itself is raised here and
   * everything past it - the plugin, the placement, the focus - is the real path.
   */
  rightClick(): void;
  /** How many rows the first table holds */
  tableRows(): number;
  /** Locks `length` characters `offset` into a block's content */
  lock(blockIndex: number, offset: number, length: number): boolean;
  /** The text every locked control in the document holds */
  lockedText(): string;
}

declare global {
  interface Window {
    /**
     * Put in place once the editor is ready, which every test waits for
     * (`../support/harness`), so it is declared as always being there.
     */
    docxHarness: DocxHarness;
  }
}
