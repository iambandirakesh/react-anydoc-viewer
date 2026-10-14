/**
 * A single slide.
 *
 * Renders the slide background, the shapes inherited from the layout/master
 * (behind), then the slide's own shapes in z-order. When a slide turns out to be
 * genuinely empty it says so instead of silently showing a white rectangle,
 * which is the difference between "empty slide" and "renderer failure".
 */

import type { CSSProperties, ReactNode } from "react";
import type { SlideModel } from "../model/types";
import type { MediaStore } from "../parser/media";
import { cssBackground } from "./style";
import { createLinkBuilder, ShapeNodeView, type ShapeRenderContext } from "./RenderShape";

export const DEFAULT_CHART_PALETTE = [
  "#4472C4",
  "#ED7D31",
  "#A5A5A5",
  "#FFC000",
  "#5B9BD5",
  "#70AD47",
  "#264478",
  "#9E480E",
  "#636363",
  "#997300",
];

export interface SlideViewProps {
  slide: SlideModel;
  media: MediaStore;
  /** natural slide size in CSS px (unscaled) */
  width: number;
  height: number;
  /** stable namespace so generated gradient/clip ids never collide */
  scope: string;
  interactive: boolean;
  palette?: string[];
  /** message shown for a slide that really has no content */
  emptyMessage?: string;
  warn?: (code: string, message: string) => void;
}

export function SlideView({
  slide,
  media,
  width,
  height,
  scope,
  interactive,
  palette = DEFAULT_CHART_PALETTE,
  emptyMessage = "This page is empty",
  warn,
}: SlideViewProps): ReactNode {
  const context: ShapeRenderContext = {
    media,
    scope,
    interactive,
    palette,
    warn: warn ?? (() => undefined),
  };
  const buildLink = createLinkBuilder(interactive);

  const background = cssBackground(slide.background) ?? "#FFFFFF";
  const containerStyle: CSSProperties = {
    position: "relative",
    width: `${width}px`,
    height: `${height}px`,
    overflow: "hidden",
    background,
    // Slides are content, not chrome: keep text crisp when scaled.
    contain: "layout paint style",
  };

  const hasContent = slide.shapes.length > 0 || slide.inheritedShapes.length > 0;
  const hasBackgroundImage = slide.background?.type === "image";

  return (
    <div
      className="radv-slide"
      style={containerStyle}
      data-slide-index={slide.index}
      data-slide-empty={hasContent ? undefined : "true"}
      role="group"
      aria-label={`Slide ${slide.index + 1}${slide.name ? `: ${slide.name}` : ""}`}
    >
      <div style={{ position: "relative", width: "100%", height: "100%" }}>
        {slide.inheritedShapes.map((node, index) => (
          <ShapeNodeView key={`i-${node.id}-${index}`} node={node} context={context} buildLink={buildLink} />
        ))}
        {slide.shapes.map((node, index) => (
          <ShapeNodeView key={`s-${node.id}-${index}`} node={node} context={context} buildLink={buildLink} />
        ))}
      </div>
      {/* Thumbnails pass an empty message so a small preview is not labelled;
          an empty `role="note"` node would still show up in the a11y tree. */}
      {!hasContent && !hasBackgroundImage && emptyMessage ? (
        <div className="radv-empty-slide" role="note">
          {emptyMessage}
        </div>
      ) : null}
    </div>
  );
}
