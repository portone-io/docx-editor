export type TabAlignment = "start" | "center" | "end" | "decimal" | "num";

export type TabLeader =
  | "none"
  | "dot"
  | "hyphen"
  | "underscore"
  | "heavy"
  | "middleDot";

/** One effective custom tab entry. Positions are points from the paragraph's leading margin. */
export interface TabStop {
  positionPt: number;
  align: TabAlignment | "bar";
  leader?: TabLeader;
}

/** A tab entry before the paragraph-property hierarchy has been resolved. */
export type TabStopDirective = TabStop | { positionPt: number; align: "clear" };

/** Composes two OOXML tab-property layers while retaining clears for a lower layer. */
export function layerTabStopDirectives(
  base: readonly TabStopDirective[],
  over: readonly TabStopDirective[]
): TabStopDirective[] {
  const stops = new Map<number, TabStopDirective>();
  for (const stop of base) {
    stops.set(stop.positionPt, stop);
  }
  for (const stop of over) {
    stops.set(stop.positionPt, stop);
  }
  return [...stops.values()].sort(
    (left, right) => left.positionPt - right.positionPt
  );
}

/** Resolves the final OOXML tab-property layers, removing cleared positions. */
export function mergeTabStops(
  base: readonly TabStopDirective[],
  over: readonly TabStopDirective[]
): TabStop[] {
  return layerTabStopDirectives(base, over).filter(
    (stop): stop is TabStop => stop.align !== "clear"
  );
}
