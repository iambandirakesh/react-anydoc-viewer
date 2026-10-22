/**
 * Standalone image adapter.
 *
 * An image is not a shape list either, so it takes the same route as a PDF: one
 * page, drawn by a page renderer. That keeps navigation, zoom, fit modes,
 * thumbnails, presentation mode and the error boundary working without teaching
 * any of them what a PNG is.
 *
 * The page box is the image's own pixel size (1px === 1/96in, the same grid the
 * rest of the model uses), so `fit="page"` fits it exactly and `fit="actual"` is
 * pixel-accurate. Dimensions are read from the file header instead of waiting
 * for a decode, which keeps parsing synchronous; a header we cannot read is not
 * fatal, because the renderer letterboxes the image into whatever box it is
 * given.
 *
 * The bytes are always delivered to the browser as an `<img src="blob:…">`.
 * That is deliberate even for SVG: browsers do not run scripts in an image, so
 * an untrusted SVG cannot reach the host page. Nothing here ever builds markup
 * from file content.
 */

import type { PresentationModel, SlideModel } from "../model/types";
import { DocumentViewerError } from "../errors";
import {
  isBrowserRenderable,
  mimeForKind,
  sniffImageKind,
  type ImageKind,
} from "./archive";
import { createEmptyMediaStore, createByteUrl, type MediaStore } from "./media";
import { DEFAULT_COLOR_MAP, createEmptyTheme } from "./theme";
import { pxToEmu } from "./units";
import { Diagnostics, DEFAULT_MAX_DIAGNOSTICS } from "./diagnostics";
import { createImagePageRenderer } from "../components/ImagePage";
import type { PageRenderer } from "../renderer/pageRenderer";

export interface ImageInfo {
  kind: ImageKind;
  /** natural size in CSS pixels, or null when the header could not be read */
  width: number | null;
  height: number | null;
}

export interface ParsedImage {
  format: "image";
  model: PresentationModel;
  media: MediaStore;
  renderPage: PageRenderer;
  dispose(): void;
}

/** Used when a header yields no usable size; a 4:3 box at 96dpi. */
const DEFAULT_SIZE = { width: 1024, height: 768 };
/** Sizes outside this range are treated as unreadable rather than honoured. */
const MAX_DIMENSION = 20_000;

function readUint32BE(bytes: Uint8Array, offset: number): number {
  return (
    ((bytes[offset] as number) << 24) |
    ((bytes[offset + 1] as number) << 16) |
    ((bytes[offset + 2] as number) << 8) |
    (bytes[offset + 3] as number)
  ) >>> 0;
}

function readUint16BE(bytes: Uint8Array, offset: number): number {
  return ((bytes[offset] as number) << 8) | (bytes[offset + 1] as number);
}

function readUint16LE(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] as number) | ((bytes[offset + 1] as number) << 8);
}

function readInt32LE(bytes: Uint8Array, offset: number): number {
  return (
    (bytes[offset] as number) |
    ((bytes[offset + 1] as number) << 8) |
    ((bytes[offset + 2] as number) << 16) |
    ((bytes[offset + 3] as number) << 24)
  );
}

function plausible(width: number, height: number): boolean {
  return (
    Number.isFinite(width) &&
    Number.isFinite(height) &&
    width > 0 &&
    height > 0 &&
    width <= MAX_DIMENSION &&
    height <= MAX_DIMENSION
  );
}

function pngSize(bytes: Uint8Array): [number, number] | null {
  if (bytes.length < 24) return null;
  return [readUint32BE(bytes, 16), readUint32BE(bytes, 20)];
}

function gifSize(bytes: Uint8Array): [number, number] | null {
  if (bytes.length < 10) return null;
  return [readUint16LE(bytes, 6), readUint16LE(bytes, 8)];
}

function bmpSize(bytes: Uint8Array): [number, number] | null {
  if (bytes.length < 26) return null;
  // A negative height means a top-down bitmap; the magnitude is the height.
  return [readInt32LE(bytes, 18), Math.abs(readInt32LE(bytes, 22))];
}

function webpSize(bytes: Uint8Array): [number, number] | null {
  if (bytes.length < 30) return null;
  const chunk = String.fromCharCode(bytes[12] as number, bytes[13] as number, bytes[14] as number, bytes[15] as number);
  if (chunk === "VP8X") {
    const width = 1 + ((bytes[24] as number) | ((bytes[25] as number) << 8) | ((bytes[26] as number) << 16));
    const height = 1 + ((bytes[27] as number) | ((bytes[28] as number) << 8) | ((bytes[29] as number) << 16));
    return [width, height];
  }
  if (chunk === "VP8 ") {
    // Lossy: the 14-bit dimensions sit after the 3-byte frame tag.
    if (bytes.length < 30) return null;
    return [
      readUint16LE(bytes, 26) & 0x3fff,
      readUint16LE(bytes, 28) & 0x3fff,
    ];
  }
  if (chunk === "VP8L") {
    const b0 = bytes[21] as number;
    const b1 = bytes[22] as number;
    const b2 = bytes[23] as number;
    const b3 = bytes[24] as number;
    return [1 + (b0 | ((b1 & 0x3f) << 8)), 1 + ((b1 >> 6) | (b2 << 2) | ((b3 & 0x0f) << 10))];
  }
  return null;
}

/**
 * JPEG has no fixed header: the size lives in whichever `SOFn` segment appears,
 * so the marker chain has to be walked. Every other segment carries a length to
 * skip over, and `0xFF` fill bytes between segments are legal.
 */
function jpegSize(bytes: Uint8Array): [number, number] | null {
  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    let marker = bytes[offset + 1] as number;
    offset += 2;
    while (marker === 0xff && offset < bytes.length) {
      marker = bytes[offset] as number;
      offset += 1;
    }
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (offset + 1 >= bytes.length) return null;
    const length = readUint16BE(bytes, offset);
    // SOF0..SOF15, excluding DHT (0xC4), JPG (0xC8) and DAC (0xCC).
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      if (offset + 7 >= bytes.length) return null;
      return [readUint16BE(bytes, offset + 5), readUint16BE(bytes, offset + 3)];
    }
    if (length < 2) return null;
    offset += length;
  }
  return null;
}

/**
 * SVG dimensions come from `width`/`height`, or from `viewBox` when those are
 * absent (very common). Percentages and CSS units are not resolvable without a
 * layout, so they are treated as absent.
 */
function svgSize(bytes: Uint8Array): [number, number] | null {
  const head = new TextDecoder("utf-8", { fatal: false }).decode(bytes.subarray(0, 4096));
  const attribute = (name: string): number | null => {
    const match = new RegExp(`${name}\\s*=\\s*["']([^"']+)["']`, "i").exec(head);
    if (!match) return null;
    const raw = (match[1] ?? "").trim();
    if (raw.endsWith("%")) return null;
    const value = Number.parseFloat(raw);
    return Number.isFinite(value) ? value : null;
  };
  if (!/<svg[\s>]/i.test(head)) return null;
  const width = attribute("width");
  const height = attribute("height");
  if (width && height) return [width, height];
  const viewBox = /viewBox\s*=\s*["']([^"']+)["']/i.exec(head);
  if (viewBox) {
    const parts = (viewBox[1] ?? "").trim().split(/[\s,]+/).map(Number);
    if (parts.length === 4 && Number.isFinite(parts[2]) && Number.isFinite(parts[3])) {
      return [parts[2] as number, parts[3] as number];
    }
  }
  return null;
}

/** Read the natural size of an image from its header, without decoding it. */
export function readImageInfo(bytes: Uint8Array): ImageInfo {
  const kind = sniffImageKind(bytes);
  let size: [number, number] | null = null;
  switch (kind) {
    case "png":
      size = pngSize(bytes);
      break;
    case "jpeg":
      size = jpegSize(bytes);
      break;
    case "gif":
      size = gifSize(bytes);
      break;
    case "bmp":
      size = bmpSize(bytes);
      break;
    case "webp":
      size = webpSize(bytes);
      break;
    case "svg":
      size = svgSize(bytes);
      break;
    default:
      size = null;
  }
  if (!size || !plausible(size[0], size[1])) {
    return { kind, width: null, height: null };
  }
  return { kind, width: Math.round(size[0]), height: Math.round(size[1]) };
}

/** Why a recognised image cannot be shown, in the words an end user needs. */
export function unsupportedImageMessage(kind: ImageKind): string {
  const names: Partial<Record<ImageKind, string>> = {
    tiff: "a TIFF image",
    emf: "an EMF metafile",
    wmf: "a WMF metafile",
    unknown: "an image",
  };
  return `This file is ${names[kind] ?? "an image"}, which this browser cannot display. Convert it to PNG, JPEG, GIF, WebP or SVG and open that instead.`;
}

/**
 * Build a one-page model for an image.
 *
 * A kind the browser cannot decode (TIFF, EMF, WMF — the vector/metafile
 * formats OOXML documents commonly embed) is refused with a specific message
 * rather than being rendered as a broken image.
 */
export function parseImageDocument(bytes: Uint8Array, options: { maxDiagnostics?: number } = {}): ParsedImage {
  const diagnostics = new Diagnostics(options.maxDiagnostics ?? DEFAULT_MAX_DIAGNOSTICS);
  const info = readImageInfo(bytes);

  if (!isBrowserRenderable(info.kind)) {
    throw new DocumentViewerError("UNSUPPORTED_FORMAT", unsupportedImageMessage(info.kind), {
      detail: `detected image kind: ${info.kind}`,
    });
  }

  const mime = mimeForKind(info.kind);
  const stored = createByteUrl(bytes, mime);
  if (!stored) {
    throw new DocumentViewerError(
      "UNSUPPORTED_INPUT",
      "The image could not be handed to the browser: this environment provides neither object URLs nor base64 encoding.",
    );
  }

  if (info.width === null || info.height === null) {
    diagnostics.add({
      severity: "info",
      code: "image.size-unknown",
      message: `The image header does not declare a size (${info.kind}), so it is shown in a default ${DEFAULT_SIZE.width}×${DEFAULT_SIZE.height} box.`,
    });
  }

  const width = info.width ?? DEFAULT_SIZE.width;
  const height = info.height ?? DEFAULT_SIZE.height;

  const page: SlideModel = {
    index: 0,
    partName: "image#1",
    name: null,
    slideshowHidden: false,
    // The image is drawn by the page renderer, not from shapes.
    shapes: [],
    inheritedShapes: [],
    background: { type: "solid", color: { r: 255, g: 255, b: 255, a: 1 } },
    diagnostics: [],
  };

  const model: PresentationModel = {
    width: pxToEmu(width),
    height: pxToEmu(height),
    slides: [page],
    theme: createEmptyTheme(),
    colorMap: DEFAULT_COLOR_MAP,
    textStyles: { title: [], body: [], other: [] },
    title: null,
    author: null,
    partCount: 1,
    diagnostics: diagnostics.list(),
  };

  const media = createEmptyMediaStore();
  const renderPage = createImagePageRenderer({ src: stored.url, label: `${width} × ${height}` });

  return {
    format: "image",
    model,
    media,
    renderPage,
    dispose: () => {
      media.dispose();
      stored.revoke();
    },
  };
}
