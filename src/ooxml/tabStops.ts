import type {
  TabAlignment,
  TabLeader,
  TabStopDirective,
} from "../model/tabStops";
import { twipsToPt, wAttr } from "./units";
import { childByLocalName } from "./xml";

const ALIGN_BY_VALUE: Record<string, TabAlignment | "bar" | "clear"> = {
  left: "start",
  start: "start",
  center: "center",
  right: "end",
  end: "end",
  decimal: "decimal",
  num: "num",
  bar: "bar",
  clear: "clear",
};

const LEADERS = new Set<TabLeader>([
  "none",
  "dot",
  "hyphen",
  "underscore",
  "heavy",
  "middleDot",
]);

function tabLeader(value: string | null): TabLeader | undefined {
  return value !== null && LEADERS.has(value as TabLeader)
    ? (value as TabLeader)
    : undefined;
}

/** Reads one `w:tabs` property without resolving the paragraph-property hierarchy. */
export function readTabStopDirectives(pPr: Element | null): TabStopDirective[] {
  if (!pPr) return [];
  const tabs = childByLocalName(pPr, "tabs");
  if (!tabs) return [];

  const stops: TabStopDirective[] = [];
  for (const tab of Array.from(tabs.children)) {
    if (tab.localName !== "tab") continue;
    const positionPt = twipsToPt(wAttr(tab, "pos"));
    const align = ALIGN_BY_VALUE[wAttr(tab, "val") ?? ""];
    if (positionPt === null || align === undefined) continue;
    if (align === "clear") {
      stops.push({ positionPt, align });
      continue;
    }
    const leader = tabLeader(wAttr(tab, "leader"));
    stops.push({
      positionPt,
      align,
      ...(leader === undefined ? {} : { leader }),
    });
  }
  return stops;
}
