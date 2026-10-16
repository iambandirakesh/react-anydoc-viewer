import { memo, type ReactNode } from "react";
import type { PresentationModel, SlideModel } from "../model/types";
import type { MediaStore } from "../parser/media";
import { SlideView } from "../renderer/SlideView";
import type { PageRenderer } from "../renderer/pageRenderer";
import { ErrorBoundary } from "./ErrorBoundary";
import type { ElementSize, MeasuredRef } from "../hooks/useElementSize";
import type { ViewerLabels } from "../types/public";

export interface SlideCanvasProps {
  model: PresentationModel;
  media: MediaStore;
  slide: SlideModel | null;
  slideIndex: number;
  slideWidth: number;
  slideHeight: number;
  scale: number;
  /** measurement of the scroll container, owned by the viewer */
  size: ElementSize;
  containerRef: MeasuredRef<HTMLDivElement>;
  labels: ViewerLabels;
  onRenderIssue?: (code: string, message: string) => void;
  onSlideRenderError?: (error: Error, index: number) => void;
  renderSlideError?: (error: Error, slideIndex: number) => ReactNode;
  emptySlideMessage?: string;
  /** draws a page whose content is not a shape list (PDF) */
  renderPage?: PageRenderer | null;
  /**
   * Space to keep free below the slide, in pixels, because the floating toolbar
   * overlaps the bottom of the stage. As a margin on the scaled slide it pushes
   * the slide up by the same amount the fit calculation already reserves.
   */
  bottomClearance?: number;
}

const MemoSlideView = memo(SlideView);

/**
 * The scrollable slide area.
 *
 * The slide is rendered at its natural size and scaled with a CSS transform, so
 * zooming never re-flows text (matching how PowerPoint zooms) and the layout
 * cost stays constant while scrolling.
 */
export function SlideCanvas({
  model,
  media,
  slide,
  slideIndex,
  slideWidth,
  slideHeight,
  scale,
  size,
  containerRef,
  labels,
  onRenderIssue,
  onSlideRenderError,
  renderSlideError,
  emptySlideMessage,
  renderPage,
  bottomClearance = 0,
}: SlideCanvasProps): ReactNode {
  const scaledWidth = slideWidth * scale;
  const scrollable = scrollableAxis(scaledWidth, size.width) || scrollableAxis(slideHeight * scale, size.height);

  return (
    <div
      className={`radv-canvas${scrollable ? " radv-canvas-scroll" : ""}`}
      ref={containerRef}
      data-measured={size.measured ? "true" : "false"}
    >
      {!size.measured ? (
        <div className="radv-notice" role="note">
          {labels.noMeasure}
        </div>
      ) : slide ? (
        <div
          className="radv-canvas-inner"
          style={{
            width: `${scaledWidth}px`,
            height: `${slideHeight * scale}px`,
            marginBottom: bottomClearance > 0 ? `${bottomClearance}px` : undefined,
          }}
        >
          <div
            style={{
              width: `${slideWidth}px`,
              height: `${slideHeight}px`,
              transform: `scale(${scale})`,
              transformOrigin: "top left",
            }}
          >
            <ErrorBoundary
              resetKey={`${model.slides.length}-${slideIndex}`}
              onError={(error) => onSlideRenderError?.(error, slideIndex)}
              fallback={(error) =>
                renderSlideError ? (
                  <>{renderSlideError(error, slideIndex)}</>
                ) : (
                  <div className="radv-slide-error" role="alert" data-testid="rv-slide-error">
                    <strong>{labels.renderError}</strong>
                    <span>
                      {labels.slideOf(slideIndex + 1, model.slides.length)} — {error.message}
                    </span>
                  </div>
                )
              }
            >
              {renderPage ? (
                renderPage({
                  page: slide,
                  width: slideWidth,
                  height: slideHeight,
                  scale,
                  interactive: true,
                  media,
                  labels,
                  warn: onRenderIssue ?? (() => undefined),
                })
              ) : (
                <MemoSlideView
                  slide={slide}
                  media={media}
                  width={slideWidth}
                  height={slideHeight}
                  scope={`slide-${slideIndex}`}
                  interactive
                  emptyMessage={emptySlideMessage ?? labels.emptySlide}
                  warn={onRenderIssue}
                />
              )}
            </ErrorBoundary>
          </div>
        </div>
      ) : (
        <div className="radv-notice" role="note">
          {labels.emptySlide}
        </div>
      )}
    </div>
  );
}

function scrollableAxis(content: number, container: number): boolean {
  return container > 0 && content > container + 1;
}
