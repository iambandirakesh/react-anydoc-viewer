import type { ReactNode } from "react";
import type { FitMode, ViewerLabels } from "../types/public";

export interface ViewerToolbarProps {
  labels: ViewerLabels;
  slideIndex: number;
  slideCount: number;
  zoom: number;
  fit: FitMode;
  thumbnailsVisible: boolean;
  allowThumbnailToggle: boolean;
  fullscreenSupported: boolean;
  isFullscreen: boolean;
  showPresentationControl: boolean;
  presentationAvailable: boolean;
  onPresent: () => void;
  canZoomOut: boolean;
  canZoomIn: boolean;
  debug: boolean;
  onPrevious: () => void;
  onNext: () => void;
  onFirst: () => void;
  onLast: () => void;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onFit: (fit: FitMode) => void;
  onToggleThumbnails: () => void;
  onToggleFullscreen: () => void;
  onRetry: () => void;
  extra?: ReactNode;
}

export function ViewerToolbar(props: ViewerToolbarProps): ReactNode {
  const {
    labels,
    slideIndex,
    slideCount,
    zoom,
    fit,
    allowThumbnailToggle,
    thumbnailsVisible,
    fullscreenSupported,
    isFullscreen,
  } = props;

  const atStart = slideIndex <= 0;
  const atEnd = slideIndex >= slideCount - 1;

  return (
    <div className="radv-toolbar" role="toolbar" aria-label={labels.viewerRegion}>
      <div className="radv-toolbar-group">
        <button
          type="button"
          className="radv-btn"
          onClick={props.onFirst}
          disabled={atStart}
          aria-label={labels.first}
          title={labels.first}
        >
          <span aria-hidden="true">«</span>
        </button>
        <button
          type="button"
          className="radv-btn"
          onClick={props.onPrevious}
          disabled={atStart}
          aria-label={labels.previous}
          title={labels.previous}
        >
          <span aria-hidden="true">‹</span>
        </button>
        <span className="radv-counter" aria-live="polite">
          {labels.slideOf(slideCount === 0 ? 0 : slideIndex + 1, slideCount)}
        </span>
        <button
          type="button"
          className="radv-btn"
          onClick={props.onNext}
          disabled={atEnd}
          aria-label={labels.next}
          title={labels.next}
        >
          <span aria-hidden="true">›</span>
        </button>
        <button
          type="button"
          className="radv-btn"
          onClick={props.onLast}
          disabled={atEnd}
          aria-label={labels.last}
          title={labels.last}
        >
          <span aria-hidden="true">»</span>
        </button>
      </div>

      <div className="radv-toolbar-group">
        <button
          type="button"
          className="radv-btn"
          onClick={props.onZoomOut}
          disabled={!props.canZoomOut}
          aria-label={labels.zoomOut}
          title={labels.zoomOut}
        >
          <span aria-hidden="true">−</span>
        </button>
        <span className="radv-zoom" aria-live="polite">
          {Math.round(zoom * 100)}%
        </span>
        <button
          type="button"
          className="radv-btn"
          onClick={props.onZoomIn}
          disabled={!props.canZoomIn}
          aria-label={labels.zoomIn}
          title={labels.zoomIn}
        >
          <span aria-hidden="true">+</span>
        </button>
        <button
          type="button"
          className={`radv-btn${fit === "page" ? " radv-btn-active" : ""}`}
          onClick={() => props.onFit("page")}
          aria-label={labels.fitPage}
          aria-pressed={fit === "page"}
          title={labels.fitPage}
        >
          <span aria-hidden="true">⤢</span>
        </button>
        <button
          type="button"
          className={`radv-btn${fit === "width" ? " radv-btn-active" : ""}`}
          onClick={() => props.onFit("width")}
          aria-label={labels.fitWidth}
          aria-pressed={fit === "width"}
          title={labels.fitWidth}
        >
          <span aria-hidden="true">↔</span>
        </button>
        <button
          type="button"
          className={`radv-btn${fit === "actual" ? " radv-btn-active" : ""}`}
          onClick={() => props.onFit("actual")}
          aria-label={labels.actualSize}
          aria-pressed={fit === "actual"}
          title={labels.actualSize}
        >
          <span aria-hidden="true">1:1</span>
        </button>
      </div>

      <div className="radv-toolbar-group radv-toolbar-end">
        {props.extra}
        {props.showPresentationControl ? (
          <button
            type="button"
            className="radv-btn radv-btn-present"
            onClick={props.onPresent}
            disabled={!props.presentationAvailable}
            aria-label={labels.present}
            title={labels.present}
          >
            <span aria-hidden="true">▶</span>
            <span className="radv-btn-label">{labels.present}</span>
          </button>
        ) : null}
        {allowThumbnailToggle ? (
          <button
            type="button"
            className={`radv-btn${thumbnailsVisible ? " radv-btn-active" : ""}`}
            onClick={props.onToggleThumbnails}
            aria-label={labels.toggleThumbnails}
            aria-pressed={thumbnailsVisible}
            title={labels.toggleThumbnails}
          >
            <span aria-hidden="true">▤</span>
          </button>
        ) : null}
        {fullscreenSupported ? (
          <button
            type="button"
            className={`radv-btn${isFullscreen ? " radv-btn-active" : ""}`}
            onClick={props.onToggleFullscreen}
            aria-label={isFullscreen ? labels.exitFullscreen : labels.fullscreen}
            aria-pressed={isFullscreen}
            title={isFullscreen ? labels.exitFullscreen : labels.fullscreen}
          >
            <span aria-hidden="true">{isFullscreen ? "⤡" : "⛶"}</span>
          </button>
        ) : null}
      </div>
    </div>
  );
}
