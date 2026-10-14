import { useEffect, useRef, useState, type ReactNode } from "react";
import type { PresentationModel } from "../model/types";
import type { MediaStore } from "../parser/media";
import { SlideView } from "../renderer/SlideView";
import type { PageRenderer } from "../renderer/pageRenderer";

export interface ThumbnailSidebarProps {
  model: PresentationModel;
  media: MediaStore;
  slideWidth: number;
  slideHeight: number;
  currentIndex: number;
  width?: number;
  labels: { thumbnails: string };
  onSelect: (index: number) => void;
  /** draws a page whose content is not a shape list (PDF) */
  renderPage?: PageRenderer | null;
}

/**
 * Thumbnails are the most expensive part of the UI, so each one is only mounted
 * once it scrolls near the viewport. In a 200-slide deck that keeps the initial
 * render proportional to what is actually visible.
 */
function LazyThumbnail({
  visible,
  children,
  width,
  height,
}: {
  visible: boolean;
  children: ReactNode;
  width: number;
  height: number;
}): ReactNode {
  return (
    <div className="radv-thumb-slot" style={{ width: `${width}px`, height: `${height}px` }}>
      {visible ? children : null}
    </div>
  );
}

export function ThumbnailSidebar({
  model,
  media,
  slideWidth,
  slideHeight,
  currentIndex,
  width = 148,
  labels,
  onSelect,
  renderPage,
}: ThumbnailSidebarProps): ReactNode {
  const containerRef = useRef<HTMLDivElement | null>(null);
  // The active thumbnail always mounts; the rest wait for the observer.
  const [visible, setVisible] = useState<Set<number>>(() => new Set());
  const scale = width / slideWidth;
  const thumbHeight = Math.round(slideHeight * scale);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return undefined;
    if (typeof IntersectionObserver === "undefined") {
      setVisible(new Set(model.slides.map((slide) => slide.index)));
      return undefined;
    }
    const observer = new IntersectionObserver(
      (entries) => {
        setVisible((previous) => {
          let changed = false;
          const next = new Set(previous);
          for (const entry of entries) {
            const index = Number((entry.target as HTMLElement).dataset.index ?? "-1");
            if (index < 0) continue;
            if (entry.isIntersecting && !next.has(index)) {
              next.add(index);
              changed = true;
            }
          }
          return changed ? next : previous;
        });
      },
      { root: container, rootMargin: "300px 0px" },
    );

    const items = container.querySelectorAll(".radv-thumb-item");
    items.forEach((item) => observer.observe(item));
    return () => observer.disconnect();
  }, [model, width]);

  // Keep the active thumbnail in view as slides change.
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;
    const active = container.querySelector<HTMLElement>(`[data-index="${currentIndex}"]`);
    active?.scrollIntoView({ block: "nearest" });
  }, [currentIndex]);

  return (
    <div className="radv-thumbnails" ref={containerRef} role="tablist" aria-label={labels.thumbnails}>
      {model.slides.map((slide) => {
        const isActive = slide.index === currentIndex;
        return (
          <button
            type="button"
            key={slide.index}
            className={`radv-thumb-item${isActive ? " radv-thumb-active" : ""}`}
            data-index={slide.index}
            role="tab"
            aria-selected={isActive}
            aria-label={`Slide ${slide.index + 1}`}
            onClick={() => onSelect(slide.index)}
          >
            <span className="radv-thumb-number">{slide.index + 1}</span>
            <span className="radv-thumb-frame" style={{ width: `${width}px`, height: `${thumbHeight}px` }}>
              <LazyThumbnail visible={visible.has(slide.index) || isActive} width={width} height={thumbHeight}>
                <span
                  className="radv-thumb-scale"
                  style={{
                    width: `${slideWidth}px`,
                    height: `${slideHeight}px`,
                    transform: `scale(${scale})`,
                  }}
                >
                  {renderPage ? (
                    renderPage({
                      page: slide,
                      width: slideWidth,
                      height: slideHeight,
                      scale,
                      interactive: false,
                      media,
                      warn: () => undefined,
                    })
                  ) : (
                    <SlideView
                      slide={slide}
                      media={media}
                      width={slideWidth}
                      height={slideHeight}
                      scope={`thumb-${slide.index}`}
                      interactive={false}
                      emptyMessage=""
                    />
                  )}
                </span>
              </LazyThumbnail>
            </span>
          </button>
        );
      })}
    </div>
  );
}
