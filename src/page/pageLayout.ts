/**
 * The arithmetic that makes one continuous sheet look like separate pages.
 *
 * The push amounts, the space opened at each page break and the places where a gap opens up
 * between pages are all decided here in one pass, and the view simply uses those values.
 * A page break is the one thing split inside a block; nothing else is broken line by line, so
 * the places where Word actually breaks may differ.
 *
 * The size of a page is the document's own (`docx/pageGeometry`), not a fixed A4: a Letter
 * document breaks at Letter's height and fits its tables to Letter's width. A document written in
 * several sections breaks each block at the height of the section that block sits in
 * (`docx/sections`), and a section opens a page of its own unless it declares `continuous`, which
 * carries it on down the page the section before it ends on.
 *
 * The width is the exception: one sheet is drawn at one width (`styles/editor.css`), which is the
 * first section's, so a section on wider paper is paginated on its own height but drawn on the
 * paper the sheet holds.
 */

import {
  A4_PORTRAIT,
  bodyHeightTwips,
  bodyWidth,
  type PageGeometry,
  twipsToPx,
} from "../docx/pageGeometry";
import type { DocumentSection, SectionProperties } from "../docx/sections";
import { editorCssVariables } from "../styles/classNames";
import type { MeasuredBlock, PageCut } from "./blockKinds";
import type { DemandBand, PageDemand } from "./demands";

export type { MeasuredBlock };

/** The paper measured in the pixels the sheet is drawn with */
export interface PagePixels {
  /** The whole sheet, padding included */
  pageWidth: number;
  pageHeight: number;
  /** The width one line of body text occupies: the sheet less its side padding */
  bodyWidth: number;
  /** The height the body occupies on one page */
  bodyHeight: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  /** From the end of one page's body to the top of the next page's body */
  pageStep: number;
}

/** The grey gap visible between one page and the next. The one number the paper does not decide */
export const PAGE_SPLIT_PX = 24;

/**
 * The paper as the sheet draws it.
 *
 * Every width and height on screen comes from here, so the document's own `w:pgSz` and
 * `w:pgMar` decide the sheet, the page breaks, and the width a table or an image is fitted to.
 */
export function pagePixels(geometry: PageGeometry): PagePixels {
  const marginTop = twipsToPx(geometry.marginTopTwips);
  const marginBottom = twipsToPx(geometry.marginBottomTwips);
  return {
    pageWidth: twipsToPx(geometry.widthTwips),
    pageHeight: twipsToPx(geometry.heightTwips),
    // The same width a table is fitted to, so the sheet and the tables on it cannot drift
    bodyWidth: bodyWidth(geometry).px,
    bodyHeight: twipsToPx(bodyHeightTwips(geometry)),
    marginLeft: twipsToPx(geometry.marginLeftTwips),
    marginRight: twipsToPx(geometry.marginRightTwips),
    marginTop,
    marginBottom,
    pageStep: marginTop + marginBottom + PAGE_SPLIT_PX,
  };
}

/** The paper a document that names none is drawn on */
export const A4_PAGE_PIXELS: PagePixels = pagePixels(A4_PORTRAIT);

/** The paper one section of a document is laid out on */
export interface SectionPixels {
  /**
   * The last position this section covers. A section is closed by the paragraph carrying its
   * `w:sectPr` (§17.6.17), so the block after that one opens the section that follows, and the
   * section the body itself closes reaches past every block.
   */
  untilPos: number;
  pixels: PagePixels;
  /**
   * How this section starts (§17.6.22). `continuous` carries it on down the page the section
   * before it ends on; every other kind opens a page.
   */
  type: SectionProperties["type"];
}

/** The one section a document naming none is laid out on */
export const A4_SECTION_PIXELS: readonly SectionPixels[] = [
  { untilPos: Number.POSITIVE_INFINITY, pixels: A4_PAGE_PIXELS, type: null },
];

/**
 * The paper of every section of a document, in the pixels the sheet is drawn with.
 *
 * Reading the sections walks every block (`docx/sections`), so this runs once per document change
 * rather than once per measurement.
 */
export function sectionPixels(
  sections: readonly DocumentSection[]
): SectionPixels[] {
  return sections.map(({ anchor, props }) => ({
    untilPos:
      anchor.kind === "paragraph" ? anchor.pos : Number.POSITIVE_INFINITY,
    pixels: pagePixels(props.geometry),
    type: props.type,
  }));
}

/** Which section the block at this position belongs to: the first one reaching that far */
function sectionIndexAt(
  sections: readonly SectionPixels[],
  pos: number
): number {
  const reaching = sections.findIndex((section) => pos <= section.untilPos);
  return reaching === -1 ? sections.length - 1 : reaching;
}

/**
 * The paper the block at this position is laid out on, which is the paper of every page that
 * block stands on. A4 where the list names no section at all.
 */
export function sectionPaperAt(
  sections: readonly SectionPixels[],
  pos: number
): PagePixels {
  return sections[sectionIndexAt(sections, pos)]?.pixels ?? A4_PAGE_PIXELS;
}

/** A pixel length as CSS writes it, rounded so the sheet and the arithmetic agree to the pixel */
function px(value: number): string {
  return `${Math.round(value * 100) / 100}px`;
}

/**
 * The paper the document names, as the CSS variables the sheet is drawn from.
 *
 * The page breaks above are worked out from the same geometry, so the divider drawn between
 * pages lands where the paper actually ends. `editor.css` carries A4 as the fallback for the
 * moment before a document is opened.
 */
export function pageGeometryStyle(page: PagePixels): string {
  return [
    `${editorCssVariables.pageWidth}:${px(page.pageWidth)}`,
    `${editorCssVariables.pageHeight}:${px(page.pageHeight)}`,
    `${editorCssVariables.pageMarginTop}:${px(page.marginTop)}`,
    `${editorCssVariables.pageMarginRight}:${px(page.marginRight)}`,
    `${editorCssVariables.pageMarginBottom}:${px(page.marginBottom)}`,
    `${editorCssVariables.pageMarginLeft}:${px(page.marginLeft)}`,
  ].join(";");
}

/**
 * The slack that keeps a page from flipping when the measured values wobble below the
 * decimal point
 */
const TOLERANCE_PX = 0.5;

/** One block to be moved down to the next page */
export interface BlockPush {
  pos: number;
  /** The gap to place above the block: the original gap plus the push */
  marginTop: number;
  /** Of that, the extra amount pushed for the page's sake */
  push: number;
}

/** A place where one page parts from the next */
export interface PageSplit {
  /** Where the previous page's body ends, measured from the top of the body */
  y: number;
  /** The number of the page that follows this place */
  page: number;
  /** Whether the document records a break at this place */
  forced: boolean;
  /**
   * Whether the block is taller than one page, so it cannot be pushed and the text
   * crosses this place
   */
  crossed: boolean;
}

/** Where one page starts */
export interface PageStart {
  page: number;
  /** Where this page's body starts, measured from the top of the body */
  bodyStart: number;
  /**
   * The position of the block this page opens with, which is the block the text crossing into it
   * belongs to. 0 for a document holding no block at all
   */
  pos: number;
  /**
   * The place this page takes within its own section, counted from 1 again at every section it
   * opens. A document read with no section table at all is one section, so this counts the whole
   * of it
   */
  pageInSection: number;
  /**
   * Whether text crosses over from the previous page, so this page continues with no top
   * margin
   */
  crossed: boolean;
}

/** The room one page keeps at the foot of its body for one band (`page/demands`) */
export interface PageReservation {
  page: number;
  band: string;
  /** The ids the page keeps, in the order it took them: those carried over from the page before first */
  ids: readonly string[];
  /**
   * Where the band starts, measured from the top of the body like a split. The bands a page keeps
   * stand one under the next in the order it took them, from the end of its body above the room.
   * A page that could not give its bands all the room they ask for gives them what lies under the
   * text that has to stay on it, so a band starts no higher than that text ends
   */
  top: number;
  /**
   * The band's overhead and the height of each id, which the page's body ends above. A demand
   * taller than an empty page's body is counted whole, though the page has less to give it
   */
  height: number;
}

export interface PageLayout {
  /** Only the blocks to be moved down to the next page */
  pushes: BlockPush[];
  /** Where the layout parted a block, and the space it opened before the continued piece */
  cuts: PageCut[];
  splits: PageSplit[];
  /** One per page, the first page included */
  pages: PageStart[];
  /** The body height with the last page filled out in full */
  bodyHeight: number;
  /**
   * The margin under that last page, which is its own section's. The sheet is padded from the
   * first section's paper alone (`editor/createEditor`), so a document ending on a deeper bottom
   * margin needs the sheet drawn down to this.
   */
  marginBottom: number;
  /** What each page keeps at its foot, page by page, one entry per band a page holds */
  reserved: readonly PageReservation[];
}

export interface PageLayoutInput {
  blocks: readonly MeasuredBlock[];
  /**
   * The paper of each section in document order (`sectionPixels`). A block is laid out on the
   * paper of the section it sits in, and the section a block opens starts a page of its own
   * unless that section is `continuous`.
   */
  sections: readonly SectionPixels[];
  /**
   * The bands the blocks' demands are kept in, by name. A demand naming a band not given here is
   * kept nowhere, so with none given the pages are the ones the blocks alone come to
   */
  bands?: ReadonlyMap<string, DemandBand>;
}

/** Where one block is laid out: the paper of its section, and whether it opens a page */
interface BlockPaper {
  paper: PagePixels;
  /**
   * Whether this block opens a section that asks for a page of its own, which is every kind of
   * section start but `continuous` (§17.6.22)
   */
  opensPage: boolean;
}

/** A page as the pass opens it, before the sections number it */
type PageOpening = Omit<PageStart, "pageInSection">;

function round(value: number): number {
  return Math.round(value * 100) / 100;
}

const NO_BANDS: ReadonlyMap<string, DemandBand> = new Map();
const NO_DEMANDS: readonly PageDemand[] = [];

/** A demand as the pass came to it, numbered in that order */
interface Taken {
  readonly demand: PageDemand;
  readonly order: number;
}

/** What the page being filled keeps at its foot */
interface Room {
  /** The ids each band keeps, in the order they were taken */
  readonly bands: Map<string, Set<string>>;
  /** What is kept, one per id of a band, in the order taken: the last taken is let go first */
  readonly kept: Taken[];
  height: number;
}

function emptyRoom(): Room {
  return { bands: new Map(), kept: [], height: 0 };
}

/** Keeps a demand once per id of its band; a demand of a band not given is kept nowhere */
function keep(
  bands: ReadonlyMap<string, DemandBand>,
  room: Room,
  taken: Taken
): void {
  const { demand } = taken;
  const band = bands.get(demand.band);
  if (!band || room.bands.get(demand.band)?.has(demand.id)) return;
  const ids = room.bands.get(demand.band);
  if (ids) {
    ids.add(demand.id);
  } else {
    room.bands.set(demand.band, new Set([demand.id]));
    room.height += band.overhead;
  }
  room.kept.push(taken);
  room.height += band.heights.get(demand.id) ?? 0;
}

/** Lets go of the demand taken last, and of its band's overhead once the band keeps nothing else */
function letGo(
  bands: ReadonlyMap<string, DemandBand>,
  room: Room
): Taken | undefined {
  const taken = room.kept.pop();
  const band = taken && bands.get(taken.demand.band);
  const ids = taken && room.bands.get(taken.demand.band);
  if (!taken || !band || !ids) return undefined;
  ids.delete(taken.demand.id);
  room.height -= band.heights.get(taken.demand.id) ?? 0;
  if (ids.size === 0) {
    room.bands.delete(taken.demand.band);
    room.height -= band.overhead;
  }
  if (room.kept.length === 0) room.height = 0;
  return taken;
}

/** The room one band takes for these ids: its overhead once, and the height of each */
function bandHeight(band: DemandBand, ids: Iterable<string>): number {
  let height = band.overhead;
  for (const id of ids) height += band.heights.get(id) ?? 0;
  return height;
}

/** How much taller a room grows by keeping these demands as well */
function roomAdded(
  bands: ReadonlyMap<string, DemandBand>,
  room: Room,
  demands: readonly PageDemand[]
): number {
  if (demands.length === 0) return 0;
  const trial: Room = {
    bands: new Map(
      Array.from(room.bands, ([name, ids]) => [name, new Set(ids)])
    ),
    kept: [...room.kept],
    height: room.height,
  };
  for (const demand of demands) keep(bands, trial, { demand, order: 0 });
  return trial.height - room.height;
}

/** A block's demands in the order their places stand in it */
function demandsInOrder(block: MeasuredBlock): readonly PageDemand[] {
  const demands = block.demands ?? NO_DEMANDS;
  return demands.length < 2
    ? demands
    : [...demands].sort((a, b) => a.offset - b.offset);
}

/**
 * The pages numbered again from one at the top of every section they open.
 *
 * Pages come in the order they are drawn and a section covers a run of blocks, so a page belongs
 * to the section the block it opens with belongs to, and a new section starts wherever that answer
 * changes. A `continuous` section opens no page of its own, so the page it starts partway down
 * stays with the section that opened that page and the count restarts only on the page after.
 */
function withinSections(
  pages: readonly PageOpening[],
  sections: readonly SectionPixels[]
): PageStart[] {
  let section: number | null = null;
  let within = 0;
  return pages.map((page) => {
    const index = sectionIndexAt(sections, page.pos);
    within = index === section ? within + 1 : 1;
    section = index;
    return { ...page, pageInSection: within };
  });
}

/**
 * Whether the keep a block asks for can hold between it and the block after it.
 * A block parted inside itself starts its last piece where the cut put it, so there is nothing a
 * keep could move; and a page the document starts between the two is one no keep can close, the
 * start of a section that opens one included.
 */
function keepsWithNext(
  block: MeasuredBlock,
  next: MeasuredBlock,
  opensPage: boolean
): boolean {
  return (
    block.keepWithNext &&
    !opensPage &&
    block.candidates.length === 0 &&
    !block.breakAfter &&
    !next.breakBefore
  );
}

/**
 * What a block kept with the next one brings onto the page it starts on, by block index: itself,
 * every block kept after it, and the first piece of the block the keeps end at (§17.3.1.14).
 *
 * A run of keeps no page can hold is let go as a whole, and none of its blocks is listed. Let go
 * from its head alone, a run longer than a page would leave a page short wherever its tail first
 * fit, for a keep that cannot be kept anyway.
 */
function keptExtents(
  blocks: readonly MeasuredBlock[],
  laid: readonly BlockPaper[]
): Map<number, number> {
  const fitsPage = (extent: number, index: number) => {
    const height = laid[index]?.paper.bodyHeight ?? 0;
    return extent <= height + TOLERANCE_PX;
  };
  const extents = new Map<number, number>();
  /** The run being walked up from its end, the block nearest its head last */
  let run: { index: number; extent: number }[] = [];
  const settle = () => {
    const head = run.at(-1);
    if (head && fitsPage(head.extent, head.index)) {
      for (const { index, extent } of run) {
        extents.set(index, extent);
      }
    }
    run = [];
  };
  for (let index = blocks.length - 2; index >= 0; index -= 1) {
    const block = blocks[index];
    const next = blocks[index + 1];
    const opensPage = laid[index + 1]?.opensPage === true;
    if (block && next && keepsWithNext(block, next, opensPage)) {
      const below = run.at(-1)?.extent ?? next.minFirstPiece;
      run.push({ index, extent: block.height + next.gap + below });
    } else {
      settle();
    }
  }
  settle();
  return extents;
}

/**
 * Distributes the blocks across pages.
 *
 * A block straddling the end of a page is pushed to the top of the next page, and the
 * blocks below it move down by the same amount.
 * A block taller than one page cannot be pushed, so it is left where it is and only the
 * places it crosses are reported.
 * A block is parted only where its measurer offered a candidate. A forced one - a page break
 * written into the text - is filled out to the end of the page it falls on, which carries the rest
 * of that block to the top of the next page while the block itself stays whole. An optional one -
 * a row boundary in a table - is taken only where the piece after it would run off the page.
 * A block kept with the next one is pushed together with what it is kept with: the blocks kept
 * after it and the first piece of the block the keeps end at.
 * Each block is measured against the paper of its own section, and the first block of a section
 * opens a page of its own unless the section it opens declares `continuous`.
 *
 * A demand (`page/demands`) asks the page its place lands on for room at the foot, and that page's
 * body ends above the room: whatever would stand in it is pushed, parted or crossed onto the next
 * page exactly as at the end of the paper, the demand's own piece included when the room it asks
 * for does not fit beneath that piece. The page ends and page tops stay where the paper puts them,
 * except the place a block taller than a page crosses, which is the end of the body above the room.
 * A demand is never parted. Where the content that has to stay on a page leaves no room for it, it
 * is carried whole to the top of the next page's band with every demand after it on that page, and
 * past the last block a page is opened for them; one taller than an empty page's body could fit
 * nowhere, so it stays where it landed. A block that opens its page is never pushed off it to make
 * room, which would leave a page holding nothing but what was carried onto it.
 */
export function pageLayout({
  blocks,
  sections,
  bands = NO_BANDS,
}: PageLayoutInput): PageLayout {
  const pushes: BlockPush[] = [];
  const cuts: PageCut[] = [];
  const splits: PageSplit[] = [];
  const reserved: PageReservation[] = [];
  /** The block being placed, which is the one a page opened along the way starts with */
  let blockPos = blocks[0]?.pos ?? 0;
  const opened: PageOpening[] = [
    { page: 1, bodyStart: 0, pos: blockPos, crossed: false },
  ];
  // A section is read into a geometry that always leaves a body to draw on
  // (`docx/pageGeometry`), so a paper with no body is a hand-built one; a list with no such paper
  // at all cannot hold a single page, let alone a boundary between two
  const roomy = sections.find(
    (section) => section.pixels.bodyHeight > 0
  )?.pixels;
  if (roomy === undefined) {
    return {
      pushes,
      cuts,
      splits,
      pages: withinSections(opened, sections),
      bodyHeight: 0,
      marginBottom: 0,
      reserved,
    };
  }

  /** The section each block sits in, in document order */
  const blockSections = blocks.map((block) =>
    sectionIndexAt(sections, block.pos)
  );
  /** Where each block is laid out, read once for the whole pass */
  const laid: BlockPaper[] = blockSections.map((section, index) => {
    const pixels = sections[section]?.pixels;
    return {
      paper: pixels && pixels.bodyHeight > 0 ? pixels : roomy,
      opensPage:
        index > 0 &&
        section !== blockSections[index - 1] &&
        sections[section]?.type !== "continuous",
    };
  });

  let pageStart = 0;
  let cursor = 0;
  /** The paper of the block being placed, which is the paper of the page it stands on */
  let paper = laid[0]?.paper ?? roomy;

  /**
   * From the end of the body of the page being closed to the top of the body of the page opening:
   * the closed page keeps its own bottom margin and the opening one its own top margin, which
   * comes to the same step on both sides of a boundary inside one section.
   */
  const stepTo = (opening: PagePixels) =>
    paper.pageStep - paper.marginTop + opening.marginTop;

  /** What the page being filled keeps at its foot */
  let room = emptyRoom();
  /** What a page could not keep, waiting for the top of the next page's band */
  let carried: Taken[] = [];
  /** How many demands the pass has come to, which numbers the next one */
  let reached = 0;
  /** How far down the page being filled the content that has to stay on it reaches */
  let held = 0;

  /** The room this page reserves, which never reaches up past the content that has to stay on it */
  const reservedHere = () =>
    Math.min(room.height, Math.max(0, pageStart + paper.bodyHeight - held));
  /** Where the body of the page being filled ends, above the room it reserves */
  const bodyEnd = () => pageStart + paper.bodyHeight - reservedHere();
  /** Whether what reaches down to `bottom` ends on the paper of the page being filled */
  const endsOnPage = (bottom: number) =>
    bottom <= pageStart + paper.bodyHeight + TOLERANCE_PX;

  /**
   * Lets go of the demands taken last until the content that has to stay on this page, down to
   * `floor`, stands above the room. A demand no empty page could keep is kept here all the same.
   */
  const settle = (floor: number, onPaper: PagePixels = paper) => {
    held = Math.max(held, floor);
    while (held + room.height > pageStart + onPaper.bodyHeight + TOLERANCE_PX) {
      const last = room.kept.at(-1);
      if (
        last === undefined ||
        roomAdded(bands, emptyRoom(), [last.demand]) >
          onPaper.bodyHeight + TOLERANCE_PX
      ) {
        return;
      }
      letGo(bands, room);
      carried.push(last);
    }
  };

  /**
   * Records what the page being closed keeps at its foot, top to bottom. The bands stand where
   * their own definitions put them (`page/demands`) rather than in the order this page's text
   * reached them, so two pages holding the same bands stack them the same way round
   */
  const recordRoom = (page: number) => {
    const held = [...room.bands]
      .flatMap(([name, ids]) => {
        const band = bands.get(name);
        return band ? [{ name, ids, band }] : [];
      })
      .sort((a, b) => b.band.order - a.band.order);
    let top = bodyEnd();
    for (const { name, ids, band } of held) {
      const height = bandHeight(band, ids);
      reserved.push({
        page,
        band: name,
        ids: [...ids],
        top: round(top),
        height: round(height),
      });
      top += height;
    }
  };

  const split = (
    y: number,
    crossed: boolean,
    forced: boolean,
    opening: PagePixels = paper
  ) => {
    const page = splits.length + 2;
    recordRoom(page - 1);
    splits.push({ y: round(y), page, forced, crossed });
    pageStart = crossed ? y : y + stepTo(opening);
    opened.push({ page, bodyStart: round(pageStart), pos: blockPos, crossed });

    room = emptyRoom();
    held = pageStart;
    const waiting = carried;
    carried = [];
    // A page lets go of what it took last first, and of a demand taken before another it let go of
    // earlier, so what waits is put back in the order the pass came to it
    waiting.sort((a, b) => a.order - b.order);
    for (const taken of waiting) keep(bands, room, taken);
    settle(pageStart, opening);
  };

  // No gap is placed along a stretch the text crosses: what a block taller than one page covers,
  // and the part of a block that runs past the end of its page before the next break in it
  const crossTo = (y: number) => {
    while (y > bodyEnd() + TOLERANCE_PX) {
      // A page whose room leaves it no body at all is crossed where its paper ends, or the pass
      // would open one empty page after another at the same place
      const end = bodyEnd();
      split(
        end > pageStart + TOLERANCE_PX ? end : pageStart + paper.bodyHeight,
        true,
        false
      );
    }
  };

  /**
   * Takes the demands of a piece reaching down to `bottom` onto the pages their places land on.
   * A piece that ends on this page's paper has to stay on it whole, and what its room does not fit
   * beneath is let go; a piece running past the paper can cross wherever the body ends, so only
   * its places have to stay above the room.
   */
  const place = (
    bottom: number,
    demands: readonly PageDemand[],
    placeOf: (demand: PageDemand) => number
  ) => {
    for (const demand of demands) {
      if (!bands.has(demand.band)) continue;
      if (!endsOnPage(bottom)) crossTo(placeOf(demand));
      const taken = { demand, order: reached };
      reached += 1;
      // Once a page has let a demand go, what stands after it on that page goes too, so no page
      // keeps a demand ahead of one standing before it
      if (carried.length > 0) {
        carried.push(taken);
      } else {
        keep(bands, room, taken);
        settle(placeOf(demand));
      }
    }
    if (endsOnPage(bottom)) settle(bottom);
    crossTo(bottom);
  };

  const kept = keptExtents(blocks, laid);
  let breakAfterPrevious = false;

  for (const [index, block] of blocks.entries()) {
    blockPos = block.pos;
    /** The paper this block is laid out on, which is the paper of the page it opens */
    const opening = laid[index]?.paper ?? paper;
    const startsSection = laid[index]?.opensPage === true;
    // The page being closed was filled with the blocks already on it, so it ends where their own
    // paper ends rather than where this block's does
    const pageEnd = pageStart + paper.bodyHeight;
    const top = cursor + block.gap;
    const startsPage =
      (block.breakBefore || breakAfterPrevious || startsSection) &&
      top > pageStart + TOLERANCE_PX;
    // Only the piece up to the first candidate has to fit on the page the block starts on, and
    // for a block kept with the next one, whatever it is kept with as well
    const first = kept.get(index) ?? block.minFirstPiece;
    const overflows = top + first > pageEnd + TOLERANCE_PX;
    const fits = first <= opening.bodyHeight + TOLERANCE_PX;
    // The demands of that piece go where it goes, so it is also pushed when the room they add does
    // not fit beneath it here, as long as an empty page could hold the piece and that room together
    const demands = demandsInOrder(block);
    const firstDemands =
      demands.length === 0
        ? NO_DEMANDS
        : demands.filter((demand) => demand.offset < first);
    const crowded =
      top > pageStart + TOLERANCE_PX &&
      top + first + roomAdded(bands, room, firstDemands) >
        bodyEnd() + TOLERANCE_PX &&
      first + roomAdded(bands, emptyRoom(), firstDemands) <=
        opening.bodyHeight + TOLERANCE_PX;

    const push =
      startsPage || (overflows && fits) || crowded
        ? Math.max(0, pageEnd + stepTo(opening) - top)
        : 0;
    if (push > 0 || startsPage) split(pageEnd, false, startsPage, opening);
    // Whether or not a page was opened for it, this block and everything after it stand on its
    // own section's paper
    paper = opening;
    if (push > 0) {
      pushes.push({
        pos: block.pos,
        marginTop: round(block.gap + push),
        push: round(push),
      });
    }

    // Offsets are read off the block with no space in it, so each cut shifts the ones after it
    const contentTop = top + push;
    const pageBodyHeight = paper.bodyHeight;
    let added = 0;
    let pieceStart = 0;
    /** The first of the block's demands no piece has taken yet */
    let untaken = 0;
    for (let index = 0; index <= block.candidates.length; index += 1) {
      const cut = index === 0 ? null : block.candidates[index - 1];
      const next = block.candidates[index];
      const pieceEnd = next?.offset ?? block.height;
      // A piece holds the places standing above the candidate it ends at, and the last piece
      // every place left
      const taken = untaken;
      while (
        untaken < demands.length &&
        (next === undefined || (demands[untaken]?.offset ?? 0) < pieceEnd)
      ) {
        untaken += 1;
      }
      const pieceDemands =
        taken === untaken ? NO_DEMANDS : demands.slice(taken, untaken);
      const pieceTop = contentTop + pieceStart + added;
      let pieceBottom = contentTop + pieceEnd + added;
      if (cut) {
        // A forced cut is answered from the page its own place stands on, however many pages
        // the piece before it crossed to get there
        if (cut.forced) crossTo(pieceTop);
        const pageBottom = pageStart + pageBodyHeight;
        const size = pieceEnd - pieceStart;
        const cutHere =
          cut.forced ||
          (pieceTop > pageStart + TOLERANCE_PX &&
            ((pieceBottom > pageBottom + TOLERANCE_PX &&
              size <= pageBodyHeight + TOLERANCE_PX) ||
              (pieceBottom + roomAdded(bands, room, pieceDemands) >
                bodyEnd() + TOLERANCE_PX &&
                size + roomAdded(bands, emptyRoom(), pieceDemands) <=
                  pageBodyHeight + TOLERANCE_PX)));
        if (cutHere) {
          split(pageBottom, false, cut.forced);
          const height = Math.max(0, pageStart - pieceTop);
          cuts.push({ at: cut.at, height: round(height) });
          added += height + cut.repeatHeight;
          pieceBottom += height + cut.repeatHeight;
        }
      }
      const shift = contentTop + added;
      place(pieceBottom, pieceDemands, (demand) => shift + demand.offset);
      pieceStart = pieceEnd;
    }

    cursor = contentTop + block.height + added;
    breakAfterPrevious = block.breakAfter;
  }

  // What the last page let go of still needs a page to be kept on
  while (carried.length > 0) {
    split(pageStart + paper.bodyHeight, false, false);
  }
  recordRoom(splits.length + 1);

  return {
    pushes,
    cuts,
    splits,
    pages: withinSections(opened, sections),
    // The last page is filled out in full on the paper of the section it opens with
    bodyHeight: round(pageStart + paper.bodyHeight),
    marginBottom: paper.marginBottom,
    reserved,
  };
}
