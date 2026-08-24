export type DocxEditorZoom = "fit-width" | number;

export const DEFAULT_ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5] as const;

const MIN_ZOOM = 0.25;
const MAX_ZOOM = 2;
const MIN_FIT_WIDTH_ZOOM = 0.5;

export function normalizeZoom(zoom: DocxEditorZoom): DocxEditorZoom {
  if (zoom === "fit-width") return zoom;
  if (!Number.isFinite(zoom)) return 1;
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, zoom));
}

export function fitWidthZoom(
  containerWidth: number,
  pageWidth: number
): number {
  if (containerWidth <= 0 || pageWidth <= 0) return 1;
  const availableWidth = Math.max(0, containerWidth - 32);
  const zoom = Math.min(1, availableWidth / pageWidth);
  return Math.max(MIN_FIT_WIDTH_ZOOM, Math.round(zoom * 100) / 100);
}
