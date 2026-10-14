import type { ReactNode } from "react";
import type { FitMode, ViewerLabels } from "../types/public";
import { ZoomMenu } from "./ZoomMenu";

export interface ViewerToolbarProps {
  labels: ViewerLabels;
  slideIndex: number;
  slideCount: number;
  zoom: number;
  fit: FitMode;
  /** true when the scale was chosen by hand, so no fit item is marked */
  explicitZoom: boolean;
  thumbnailsVisible: boolean;
  allowThumbnailToggle: boolean;
  isFullscreen: boolean;
  showPresentationControl: boolean;
  presentationAvailable: boolean;
  canZoomOut: boolean;
  canZoomIn: boolean;
  onPrevious: () => void;
  onNext: () => void;
  onZoom: (scale: number) => void;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onFit: (fit: FitMode) => void;
  onToggleThumbnails: () => void;
  onToggleFullscreen: () => void;
  onPresent: () => void;
  extra?: ReactNode;
}

/**
 * The viewer chrome: a small floating pill centred at the bottom of the stage, so
 * the slide keeps the full width instead of losing a strip to a header bar. Zoom
 * and fit share one dropdown (see {@link ZoomMenu}) rather than a button per mode.
 *
 * The pill is shown while the viewer is hovered or one of its controls has focus
 * (see the stylesheet), so it stays out of the way while reading. Everything in it
 * is short on purpose: jumping to the first or last page stays available through
 * the keyboard (`Home`/`End`) and {@link DocumentViewerHandle} rather than costing
 * the pill two more buttons in every reader's view.
 */
export function ViewerToolbar(props: ViewerToolbarProps): ReactNode {
  const {
    labels,
    slideIndex,
    slideCount,
    thumbnailsVisible,
    isFullscreen,
    allowThumbnailToggle,
  } = props;

  const atStart = slideIndex <= 0;
  const atEnd = slideIndex >= slideCount - 1;

  return (
    <div className="radv-toolbar" role="toolbar" aria-label={labels.viewerRegion}>
      <div className="radv-toolbar-group">
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
      </div>

      <span className="radv-toolbar-divider" aria-hidden="true" />

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
        <ZoomMenu
          labels={labels}
          zoom={props.zoom}
          fit={props.fit}
          explicitZoom={props.explicitZoom}
          onZoom={props.onZoom}
          onFit={props.onFit}
        />
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
      </div>

      <span className="radv-toolbar-divider" aria-hidden="true" />

      <div className="radv-toolbar-group radv-toolbar-end">
        {props.extra}
        {/*
         * Icon only: the pill has to stay narrow, and the full "present from this
         * slide" text is too long for it. It remains the accessible name and the
         * tooltip, so a keyboard or screen-reader user still hears it in full.
         */}
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
      </div>
    </div>
  );
}
