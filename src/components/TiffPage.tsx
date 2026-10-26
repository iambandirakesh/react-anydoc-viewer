/**
 * TIFF page rendering.
 *
 * Browsers do not decode TIFF in an `<img>`, so the page is decoded to RGBA and
 * painted into a canvas. `utif` is imported lazily: a viewer that never opens a
 * TIFF never downloads the decoder. The decoded bitmap is contained inside its
 * page box, centred, and the canvas backing store follows the viewer's scale so
 * zooming does not simply magnify a fixed bitmap.
 *
 * The bytes are untrusted, so a failed decode becomes a labelled notice and a
 * diagnostic instead of an unhandled rejection or a blank page.
 */

import { useEffect, useRef, useState, type ReactNode } from "react";
import type { PageRenderer } from "../renderer/pageRenderer";

/** Ceiling on raster density, matching the PDF renderer. */
const MAX_DENSITY = 3;
/** Ceiling on the canvas backing store, so a hostile page cannot exhaust memory. */
const MAX_CANVAS_PIXELS = 8_000_000;

function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value;
}

function asArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  // Copy the exact slice: a view into a larger buffer must not hand UTIF the
  // neighbouring bytes of an unrelated part.
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

interface UtifModule {
  decode(buffer: ArrayBuffer): unknown[];
  decodeImage(buffer: ArrayBuffer, ifd: unknown): void;
  toRGBA8(ifd: unknown): Uint8Array;
}

function isUtifModule(value: unknown): value is UtifModule {
  if (typeof value !== "object" || value === null) return false;
  const candidate = value as Record<string, unknown>;
  return (
    typeof candidate.decode === "function" &&
    typeof candidate.decodeImage === "function" &&
    typeof candidate.toRGBA8 === "function"
  );
}

/** UTIF is CommonJS; Vite may expose it as a default export or the namespace. */
async function loadUtif(): Promise<UtifModule> {
  const imported = (await import("utif")) as unknown;
  if (isUtifModule(imported)) return imported;
  if (typeof imported === "object" && imported !== null && "default" in imported) {
    const fallback = (imported as { default?: unknown }).default;
    if (isUtifModule(fallback)) return fallback;
  }
  throw new Error("the TIFF decoder did not expose the expected API");
}

export interface TiffPageProps {
  bytes: Uint8Array;
  pageIndex: number;
  width: number;
  height: number;
  scale: number;
  onIssue?: (code: string, message: string) => void;
}

export function TiffPage({
  bytes,
  pageIndex,
  width,
  height,
  scale,
  onIssue,
}: TiffPageProps): ReactNode {
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [failed, setFailed] = useState<string | null>(null);
  const [cssSize, setCssSize] = useState({ width, height });
  const issueRef = useRef(onIssue);
  issueRef.current = onIssue;

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return undefined;
    let cancelled = false;
    setFailed(null);

    const run = async () => {
      try {
        const Utif = await loadUtif();
        if (cancelled) return;
        const buffer = asArrayBuffer(bytes);
        const directories = Utif.decode(buffer);
        const directory = directories[pageIndex];
        if (!directory) throw new Error(`the TIFF has no page ${pageIndex + 1}`);
        Utif.decodeImage(buffer, directory);
        const rgba = Utif.toRGBA8(directory);
        if (cancelled) return;

        const image = directory as { width?: unknown; height?: unknown };
        const sourceWidth = Number(image.width);
        const sourceHeight = Number(image.height);
        if (
          !Number.isFinite(sourceWidth) ||
          !Number.isFinite(sourceHeight) ||
          sourceWidth <= 0 ||
          sourceHeight <= 0
        ) {
          throw new Error("the TIFF page has no usable dimensions");
        }

        const fit = Math.min(width / sourceWidth, height / sourceHeight) || 1;
        const fitted = { width: sourceWidth * fit, height: sourceHeight * fit };
        let density = clamp(scale * (typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1), 0.5, MAX_DENSITY);
        const pixels = fitted.width * fitted.height * density * density;
        if (pixels > MAX_CANVAS_PIXELS) {
          density = Math.sqrt(MAX_CANVAS_PIXELS / (fitted.width * fitted.height));
        }

        const source = document.createElement("canvas");
        source.width = sourceWidth;
        source.height = sourceHeight;
        const sourceContext = source.getContext("2d");
        if (!sourceContext) throw new Error("this browser did not provide a 2D canvas context");
        const data = new ImageData(new Uint8ClampedArray(rgba), sourceWidth, sourceHeight);
        sourceContext.putImageData(data, 0, 0);

        canvas.width = Math.max(1, Math.floor(fitted.width * density));
        canvas.height = Math.max(1, Math.floor(fitted.height * density));
        const context = canvas.getContext("2d");
        if (!context) throw new Error("this browser did not provide a 2D canvas context");
        context.fillStyle = "#ffffff";
        context.fillRect(0, 0, canvas.width, canvas.height);
        context.imageSmoothingEnabled = true;
        context.drawImage(source, 0, 0, canvas.width, canvas.height);
        setCssSize(fitted);
      } catch (error) {
        if (cancelled) return;
        const message = error instanceof Error ? error.message : String(error);
        setFailed(message);
        issueRef.current?.("tiff.page", `TIFF page ${pageIndex + 1} could not be rendered: ${message}`);
      }
    };

    void run();
    return () => {
      cancelled = true;
    };
  }, [bytes, pageIndex, width, height, scale]);

  return (
    <div
      className="radv-tiff-page"
      data-page={pageIndex}
      style={{ width: `${width}px`, height: `${height}px` }}
    >
      {failed ? (
        <div className="radv-notice" role="note" data-testid="rv-tiff-error">
          {`TIFF page ${pageIndex + 1} could not be rendered`}
        </div>
      ) : (
        <canvas
          ref={canvasRef}
          data-testid="rv-tiff-page"
          role="img"
          aria-label={pageIndex === 0 ? "TIFF page 1" : `TIFF page ${pageIndex + 1}`}
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

/** Build the page renderer the viewer calls for a TIFF document. */
export function createTiffPageRenderer(bytes: Uint8Array): PageRenderer {
  return ({ page, width, height, scale, warn }) => (
    <TiffPage
      bytes={bytes}
      pageIndex={page.index}
      width={width}
      height={height}
      scale={scale}
      onIssue={warn}
    />
  );
}
