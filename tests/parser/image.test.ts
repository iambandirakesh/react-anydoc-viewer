import { describe, expect, it } from "vitest";
import { readImageInfo, parseImageDocument, unsupportedImageMessage } from "../../src/parser/image";
import { detectFormat, parseDocumentBytes } from "../../src/parser/document";
import { loadDocument } from "../../src/parser/document";
import { mergeLabels } from "../../src/components/labels";

/* --------------------------------------------------------------- fixtures */

/** A minimal but structurally valid PNG header: IHDR carries the size. */
function pngHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(33);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  const view = new DataView(bytes.buffer);
  view.setUint32(8, 13); // IHDR length
  bytes.set([0x49, 0x48, 0x44, 0x52], 12); // "IHDR"
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
}

function gifHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(14);
  bytes.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61], 0); // GIF89a
  const view = new DataView(bytes.buffer);
  view.setUint16(6, width, true);
  view.setUint16(8, height, true);
  return bytes;
}

/** A JPEG whose SOF0 segment declares the size (marker walk, no pixel data). */
function jpegHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array([
    0xff, 0xd8, // SOI
    0xff, 0xe0, 0x00, 0x10, // APP0, length 16
    ...new Array(14).fill(0),
    0xff, 0xc0, 0x00, 0x11, 0x08, // SOF0, length, precision
    (height >> 8) & 0xff, height & 0xff,
    (width >> 8) & 0xff, width & 0xff,
    0x03, 0x01, 0x11, 0x00, 0x02, 0x11, 0x00, 0x03, 0x11, 0x00,
  ]);
  return bytes;
}

function bmpHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(26);
  bytes.set([0x42, 0x4d], 0); // "BM"
  const view = new DataView(bytes.buffer);
  view.setInt32(18, width, true);
  view.setInt32(22, height, true);
  return bytes;
}

const SVG = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 150"><rect width="300" height="150"/></svg>');

/** A TIFF header with no readable directory; still a real image to detect. */
const TIFF = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 0x08, 0x00, 0x00, 0x00]);

/** An EMF header: a real metafile, but not one a browser draws. */
function emfHeader(): Uint8Array {
  const bytes = new Uint8Array(48);
  bytes.set([0x01, 0x00, 0x00, 0x00], 0);
  bytes.set([0x20, 0x45, 0x4d, 0x46], 40); // " EMF"
  return bytes;
}

/* ------------------------------------------------------------------ tests */

describe("readImageInfo", () => {
  it("reads the natural size from each header it supports", () => {
    expect(readImageInfo(pngHeader(800, 600))).toEqual({ kind: "png", width: 800, height: 600 });
    expect(readImageInfo(gifHeader(64, 32))).toEqual({ kind: "gif", width: 64, height: 32 });
    expect(readImageInfo(jpegHeader(1200, 900))).toEqual({ kind: "jpeg", width: 1200, height: 900 });
    expect(readImageInfo(bmpHeader(120, 90))).toEqual({ kind: "bmp", width: 120, height: 90 });
  });

  it("uses viewBox for an SVG without width/height", () => {
    expect(readImageInfo(SVG)).toEqual({ kind: "svg", width: 300, height: 150 });
  });

  it("reports an unreadable size as null instead of inventing one", () => {
    const truncated = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    expect(readImageInfo(truncated)).toEqual({ kind: "png", width: null, height: null });
    expect(readImageInfo(new Uint8Array([1, 2, 3, 4]))).toEqual({ kind: "unknown", width: null, height: null });
  });

  it("absorbs a top-down bitmap's negative height", () => {
    expect(readImageInfo(bmpHeader(10, -20))).toEqual({ kind: "bmp", width: 10, height: 20 });
  });
});

describe("detectFormat", () => {
  it("recognises an image from its bytes, not its name", () => {
    expect(detectFormat(pngHeader(1, 1))).toBe("image");
    expect(detectFormat(SVG)).toBe("image");
    // A container signature still wins over image magic bytes.
    expect(detectFormat(new TextEncoder().encode("%PDF-1.7"))).toBe("pdf");
  });

  it("does not claim unknown bytes", () => {
    expect(detectFormat(new Uint8Array([1, 2, 3, 4, 5]))).toBe("unknown");
    // A TIFF has a real signature even when its directory is damaged.
    expect(detectFormat(TIFF)).toBe("image");
  });
});

describe("parseImageDocument", () => {
  it("builds a one-page model sized from the image", async () => {
    const parsed = parseImageDocument(pngHeader(800, 600));
    expect(parsed.format).toBe("image");
    expect(parsed.model.slides).toHaveLength(1);
    // 800px at 96dpi === 800/96 inch, in EMU.
    expect(parsed.model.width).toBe(Math.round((800 / 96) * 914_400));
    expect(parsed.model.height).toBe(Math.round((600 / 96) * 914_400));
    parsed.dispose();
  });

  it("falls back to a default box and says so when the header has no size", async () => {
    const truncated = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const parsed = parseImageDocument(truncated);
    expect(parsed.model.width).toBe(Math.round((1024 / 96) * 914_400));
    expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain("image.size-unknown");
    parsed.dispose();
  });

  it("refuses a format the browser cannot decode, naming it", () => {
    expect(() => parseImageDocument(emfHeader())).toThrowError(
      expect.objectContaining({ code: "UNSUPPORTED_FORMAT" }),
    );
    expect(unsupportedImageMessage("emf")).toContain("EMF");
  });

  it("hands the bytes to the browser as an image URL", async () => {
    const parsed = await loadDocument(pngHeader(4, 4));
    expect(parsed.format).toBe("image");
    const element = parsed.renderPage?.({
      page: parsed.model.slides[0]!,
      width: 100,
      height: 100,
      scale: 1,
      interactive: true,
      media: parsed.media,
      labels: mergeLabels(undefined),
      warn: () => undefined,
    }) as { props: { src: string } } | null;
    expect(element?.props.src).toMatch(/^(blob:|data:image\/png)/);
    parsed.dispose();
  });

  it("is reachable through the async document entry point", async () => {
    const parsed = await parseDocumentBytes(SVG);
    expect(parsed.format).toBe("image");
    expect(parsed.renderPage).toBeTypeOf("function");
    parsed.dispose();
  });
});
