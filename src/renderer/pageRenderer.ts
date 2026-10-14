/**
 * The per-page renderer seam.
 *
 * A `.pptx` and a `.docx` both become pages of shapes, and the slide renderer
 * draws them. A PDF page cannot: it is a program for an engine, not a shape list.
 * Such a format supplies its own renderer instead, and the canvas and thumbnail
 * components call it for whichever page they are drawing. Everything around it —
 * navigation, zoom, presentation mode, keyboard handling, the error boundary —
 * stays format-agnostic.
 */

import type { ReactNode } from "react";
import type { SlideModel } from "../model/types";
import type { MediaStore } from "../parser/media";

export interface PageRenderContext {
  /** the page to draw; `page.index` is 0-based */
  page: SlideModel;
  /** natural page size in CSS px, before the viewer's scale is applied */
  width: number;
  height: number;
  /** effective display scale: raster output should match it to stay sharp */
  scale: number;
  /** false in thumbnails: no links, no pointer interaction */
  interactive: boolean;
  media: MediaStore;
  /** report a problem; the viewer buffers and exposes these as diagnostics */
  warn: (code: string, message: string) => void;
}

export type PageRenderer = (context: PageRenderContext) => ReactNode;
