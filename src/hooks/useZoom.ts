import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { FitMode } from "../types/public";

export const MIN_ZOOM = 0.1;
export const MAX_ZOOM = 8;
const ZOOM_LADDER = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 2, 3, 4];

function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value;
}

export interface UseZoomOptions {
  /** the `fit` prop; the user can change the mode, and a new prop value wins again */
  fit: FitMode;
  /** the `zoom` prop, when the host controls the scale */
  controlled: number | undefined;
  /** the `defaultZoom` prop; `null`/absent means "follow the fit mode" */
  defaultZoom: number | undefined;
  /** the stage box, and whether it has been measured yet */
  stage: { width: number; height: number; measured: boolean };
  /** natural page size in CSS px */
  pageWidth: number;
  pageHeight: number;
  /** space kept clear around the page on the sides/top, and at the bottom */
  inset: number;
  bottomInset: number;
  /** pages that are not fitted (worksheets) stay at 100% unless zoomed */
  fitDisabled: boolean;
  onZoomChange?: (zoom: number) => void;
  onFitChange?: (fit: FitMode) => void;
}

export interface Zoom {
  /** the effective display scale */
  scale: number;
  fit: FitMode;
  /** the explicit zoom, or `null` while the fit mode decides the scale */
  explicit: number | null;
  setZoom: (zoom: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  /** back to the fit mode's scale */
  resetZoom: () => void;
  changeFit: (fit: FitMode) => void;
}

/** Fit modes, explicit zoom and the zoom ladder used by the toolbar and keys. */
export function useZoom({
  fit,
  controlled,
  defaultZoom,
  stage,
  pageWidth,
  pageHeight,
  inset,
  bottomInset,
  fitDisabled,
  onZoomChange,
  onFitChange,
}: UseZoomOptions): Zoom {
  const [fitState, setFitState] = useState<FitMode>(fit);
  useEffect(() => setFitState(fit), [fit]);

  const [explicit, setExplicit] = useState<number | null>(defaultZoom ?? null);
  useEffect(() => {
    if (controlled !== undefined) setExplicit(controlled);
  }, [controlled]);

  const fitScale = useMemo(() => {
    if (!stage.measured) return 1;
    const availableWidth = Math.max(stage.width - inset * 2, 1);
    const availableHeight = Math.max(stage.height - inset - bottomInset, 1);
    switch (fitState) {
      case "width":
        return availableWidth / pageWidth;
      case "height":
        return availableHeight / pageHeight;
      case "page":
        return Math.min(availableWidth / pageWidth, availableHeight / pageHeight);
      default:
        return 1;
    }
  }, [bottomInset, fitState, inset, pageHeight, pageWidth, stage.height, stage.measured, stage.width]);

  const scale = explicit ?? (fitDisabled ? 1 : fitScale);
  const onZoomChangeRef = useRef(onZoomChange);
  onZoomChangeRef.current = onZoomChange;
  const onFitChangeRef = useRef(onFitChange);
  onFitChangeRef.current = onFitChange;

  const setZoom = useCallback(
    (next: number) => {
      const clamped = clamp(next, MIN_ZOOM, MAX_ZOOM);
      if (controlled === undefined) setExplicit(clamped);
      onZoomChangeRef.current?.(clamped);
    },
    [controlled],
  );

  const zoomIn = useCallback(() => {
    const next = ZOOM_LADDER.find((step) => step > scale + 0.001) ?? ZOOM_LADDER[ZOOM_LADDER.length - 1]!;
    setZoom(next);
  }, [scale, setZoom]);

  const zoomOut = useCallback(() => {
    const previous = [...ZOOM_LADDER].reverse().find((step) => step < scale - 0.001) ?? ZOOM_LADDER[0]!;
    setZoom(previous);
  }, [scale, setZoom]);

  const resetZoom = useCallback(() => {
    if (controlled === undefined) setExplicit(null);
    onZoomChangeRef.current?.(fitDisabled ? 1 : fitScale);
  }, [controlled, fitScale, fitDisabled]);

  const changeFit = useCallback(
    (next: FitMode) => {
      setFitState(next);
      if (controlled === undefined) setExplicit(null);
      onFitChangeRef.current?.(next);
    },
    [controlled],
  );

  return { scale, fit: fitState, explicit, setZoom, zoomIn, zoomOut, resetZoom, changeFit };
}
