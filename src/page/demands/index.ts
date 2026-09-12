/**
 * Room a place inside a block asks the page it lands on to keep at the foot of its body.
 *
 * A block has one kind (`page/blockKinds`), and a place asking for room can stand inside a block of
 * any kind: a paragraph, or a paragraph in a table cell. So what asks is not a kind but a source,
 * and every registered source is asked about every block whatever its kind. The layout
 * (`page/pageLayout`) knows a demand by its band alone, never by what put it there.
 */

import type { MeasureTarget } from "../blockKinds";
import { footnoteDemands } from "./footnoteDemands";

/** Room one place in a block asks the page it lands on to keep at the foot */
export interface PageDemand {
  /**
   * The place's top from the block's natural top. A source reads it off the sheet as drawn, the
   * way a kind reads a candidate (`target.sheetY(viewportY) - target.top`), and the measurement
   * takes back off the spaces the engine opened inside the block above that place
   * (`page/measureBlocks`), so no source has to know what any kind draws
   */
  readonly offset: number;
  /** What is kept. One id is kept once on a page, however many places on it ask for it */
  readonly id: string;
  /** The band at the foot of the page that keeps it (`DemandBand`) */
  readonly band: string;
}

export interface DemandSource {
  readonly name: string;
  demandsIn(target: MeasureTarget): readonly PageDemand[];
}

export interface DemandBand {
  /**
   * Where this band stands among the bands one page keeps, counted up from the foot of the body:
   * 0 stands at the very foot and a higher number above it, between the text and the band below.
   * Without it a page would stack its bands in the order its text happened to reach them, so the
   * same two bands could come out one way round on one page and the other way round on the next.
   * Two bands a page may hold together name two different places
   */
  readonly order: number;
  /** Height a page adds once when it holds any demand of this band */
  readonly overhead: number;
  /** The height each id takes in the band. An id not measured yet takes none */
  readonly heights: ReadonlyMap<string, number>;
}

/** The sources the editor asks about every block, in the order their demands are listed */
export const DEFAULT_DEMAND_SOURCES: readonly DemandSource[] = [
  footnoteDemands,
];
