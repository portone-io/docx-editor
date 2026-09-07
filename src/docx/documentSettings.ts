import { ST_TwipsMeasure } from "../ooxml/simpleTypes";
import { twipsToPt, wAttr } from "../ooxml/units";
import { childByLocalName } from "../ooxml/xml";

/** The interval OOXML assigns when settings.xml does not declare one. */
export const DEFAULT_TAB_STOP_PT = 36;

/** Reads the document-wide automatic tab interval from settings.xml. */
export function readDefaultTabStop(settings: Document | null): number | null {
  if (!settings) return null;
  const setting = childByLocalName(settings.documentElement, "defaultTabStop");
  const value = setting
    ? twipsToPt(ST_TwipsMeasure.parse(wAttr(setting, "val")))
    : null;
  // An interval of nothing would put every automatic stop in the same place
  return value !== null && value > 0 ? value : null;
}
