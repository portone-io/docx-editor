import type { TabStopDirective } from "../model/tabStops";
import { ST_SignedTwipsMeasure, ST_TabJc, ST_TabTlc } from "./simpleTypes";
import { twipsToPt, wAttr } from "./units";
import { childByLocalName } from "./xml";

/** Reads one `w:tabs` property without resolving the paragraph-property hierarchy. */
export function readTabStopDirectives(pPr: Element | null): TabStopDirective[] {
  if (!pPr) return [];
  const tabs = childByLocalName(pPr, "tabs");
  if (!tabs) return [];

  const stops: TabStopDirective[] = [];
  for (const tab of Array.from(tabs.children)) {
    if (tab.localName !== "tab") continue;
    const positionPt = twipsToPt(
      ST_SignedTwipsMeasure.parse(wAttr(tab, "pos"))
    );
    const align = ST_TabJc.parse(wAttr(tab, "val"));
    if (positionPt === null || align === null) continue;
    if (align === "clear") {
      stops.push({ positionPt, align });
      continue;
    }
    const leader = ST_TabTlc.parse(wAttr(tab, "leader"));
    stops.push({
      positionPt,
      align,
      ...(leader === null ? {} : { leader }),
    });
  }
  return stops;
}
