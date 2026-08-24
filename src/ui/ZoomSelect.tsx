import { editorClassNames } from "../styles/classNames";
import { Tooltip } from "./Tooltip";
import {
  DEFAULT_ZOOM_LEVELS,
  type DocxEditorZoom,
  normalizeZoom,
} from "./zoom";

function optionValue(zoom: DocxEditorZoom): string {
  return zoom === "fit-width" ? zoom : String(zoom);
}

function percentage(zoom: number): string {
  return `${Math.round(zoom * 100)}%`;
}

export interface ZoomSelectProps {
  zoom: DocxEditorZoom;
  onZoomChange: (zoom: DocxEditorZoom) => void;
}

export function ZoomSelect({ zoom, onZoomChange }: ZoomSelectProps) {
  const normalized = normalizeZoom(zoom);
  const levels =
    typeof normalized === "number" &&
    !DEFAULT_ZOOM_LEVELS.includes(
      normalized as (typeof DEFAULT_ZOOM_LEVELS)[number]
    )
      ? [...DEFAULT_ZOOM_LEVELS, normalized].sort((left, right) => left - right)
      : DEFAULT_ZOOM_LEVELS;

  return (
    <Tooltip label="Zoom">
      <select
        className={`${editorClassNames.toolbarSelect} ${editorClassNames.zoomSelect}`}
        aria-label="Zoom"
        value={optionValue(normalized)}
        onChange={(event) => {
          const value = event.currentTarget.value;
          onZoomChange(
            value === "fit-width" ? value : normalizeZoom(Number(value))
          );
        }}
      >
        <option value="fit-width">Fit</option>
        {levels.map((level) => (
          <option key={level} value={level}>
            {percentage(level)}
          </option>
        ))}
      </select>
    </Tooltip>
  );
}
