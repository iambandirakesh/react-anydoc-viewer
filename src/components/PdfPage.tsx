/**
 * PDF page rendering.
 *
 * pdf.js rasterises one page at a time into a canvas. The canvas backing store is
 * sized from the viewer's current scale (times the device pixel ratio), so a page
 * stays sharp while zooming instead of being upscaled from a fixed bitmap, and
 * only the visible page and the visible thumbnails are ever rendered.
 */

import { useEffect, useRef, useState, type ReactNode } from "react";
import type { PageRenderer } from "../renderer/pageRenderer";

/** Structurally typed so this module never imports pdf.js at runtime. */
export interface PdfPageLike {
  getViewport(params: { scale: number }): { width: number; height: number };
  render(params: {
    canvasContext: CanvasRenderingContext2D;
    viewport: unknown;
    canvas: HTMLCanvasElement;
  }): { promise: Promise<void>; cancel(): void };
  cleanup?(): void;
}

export interface PdfDocumentLike {
  numPages: number;
  getPage(pageNumber: number): Promise<PdfPageLike>;
}

/** Ceiling on the raster density: 3x a CSS pixel is indistinguishable at 100%. */
const MAX_DENSITY = 3;
/** Ceiling on the backing store, so a huge page cannot exhaust canvas memory. */
const MAX_CANVAS_PIXELS = 8_000_000;

function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value;
}

function isCancellation(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    (error as { name?: string }).name === "RenderingCancelledException"
  );
}

export interface PdfPageProps {
  document: PdfDocumentLike;
  pageIndex: number;
  /** natural page size in CSS px (the box the page is drawn into) */
  width: number;
  height: number;
  scale: number;
  onIssue?: (code: string, message: string) => void;
}

export function PdfPage({
  document,
  pageIndex,
  width,
  height,
  scale,
  onIssue,
}: PdfPageProps): ReactNode {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [failed, setFailed] = useState<string | null>(null);
  const [cssSize, setCssSize] = useState<{ width: number; height: number }>({ width, height });
  const issueRef = useRef(onIssue);
  issueRef.current = onIssue;

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return undefined;
    let cancelled = false;
    let task: { promise: Promise<void>; cancel(): void } | null = null;
    setFailed(null);

    const run = async () => {
      try {
        const page = await document.getPage(pageIndex + 1);
        if (cancelled) return;
        const base = page.getViewport({ scale: 1 });
        const fit = Math.min(width / base.width, height / base.height) || 1;
        const devicePixelRatio = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;

        let density = clamp(scale * devicePixelRatio, 0.5, MAX_DENSITY);
        const fitted = { width: base.width * fit, height: base.height * fit };
        const pixels = fitted.width * fitted.height * density * density;
        if (pixels > MAX_CANVAS_PIXELS) density = Math.sqrt(MAX_CANVAS_PIXELS / (fitted.width * fitted.height));

        const viewport = page.getViewport({ scale: fit * density });
        canvas.width = Math.max(1, Math.floor(viewport.width));
        canvas.height = Math.max(1, Math.floor(viewport.height));
        setCssSize({ width: fitted.width, height: fitted.height });

        const context = canvas.getContext("2d");
        if (!context) throw new Error("this browser did not provide a 2D canvas context");
        task = page.render({ canvasContext: context, viewport, canvas });
        await task.promise;
      } catch (error) {
        if (cancelled || isCancellation(error)) return;
        const message = error instanceof Error ? error.message : String(error);
        setFailed(message);
        issueRef.current?.("pdf.page", `Page ${pageIndex + 1} could not be rendered: ${message}`);
      }
    };

    void run();
    return () => {
      cancelled = true;
      task?.cancel();
    };
  }, [document, pageIndex, width, height, scale]);

  return (
    <div className="radv-pdf-page" data-page={pageIndex} style={{ width: `${width}px`, height: `${height}px` }}>
      {failed ? (
        <div className="radv-notice" role="note" data-testid="rv-pdf-error">
          {`Page ${pageIndex + 1} could not be rendered`}
        </div>
      ) : (
        <canvas
          ref={canvasRef}
          data-testid="rv-pdf-page"
          role="img"
          aria-label={`Page ${pageIndex + 1}`}
          style={{
            width: `${cssSize.width}px`,
            height: `${cssSize.height}px`,
            display: "block",
            margin: "auto",
          }}
        />
      )}
    </div>
  );
}

/** Build the page renderer the viewer calls for a loaded PDF. */
export function createPdfPageRenderer(document: PdfDocumentLike): PageRenderer {
  return ({ page, width, height, scale, warn }) => (
    <PdfPage
      document={document}
      pageIndex={page.index}
      width={width}
      height={height}
      scale={scale}
      onIssue={warn}
    />
  );
}
