import { ST_TwipsMeasure } from "../ooxml/simpleTypes";
import { isOn, twipsToPt, wAttr } from "../ooxml/units";
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

/** The compatibility settings (`w:compat`) the display reads */
export interface CompatSettings {
  /** Part 4 §14.8.3.31: no custom tab stop is created at a hanging indent */
  noTabHangInd: boolean;
}

export const NO_COMPAT: CompatSettings = { noTabHangInd: false };

export function readCompatSettings(settings: Document | null): CompatSettings {
  const compat = settings
    ? childByLocalName(settings.documentElement, "compat")
    : null;
  return { noTabHangInd: compat !== null && isOn(compat, "noTabHangInd") };
}
