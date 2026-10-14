/**
 * Source normalisation and the single loading entry point.
 *
 * Everything the viewer accepts (File, Blob, ArrayBuffer, Uint8Array, URL or a
 * string URL) funnels through here, so the viewer component only ever deals with
 * either a parsed presentation or a normalised {@link DocumentViewerError}.
 */

import { DocumentViewerError, isAbortError } from "../errors";
import { ArchiveError, readArchive, type ArchiveLimits } from "./archive";
import { XmlError } from "./xml";
import { parsePresentationArchive, type ParsedPresentation } from "./presentation";
import type { PdfJsLoader } from "./pdf";
import type { ParagraphMeasurer } from "./docx/model";

export type ViewerSource = File | Blob | ArrayBuffer | Uint8Array | string | URL;

export interface LoadOptions {
  signal?: AbortSignal;
  /** passed to `fetch` when loading from a URL */
  fetchOptions?: RequestInit;
  limits?: Partial<ArchiveLimits>;
  maxDiagnostics?: number;
  /** maximum number of bytes accepted from a URL response (default 128MiB) */
  maxDownloadBytes?: number;
  /** URL of pdf.js's worker module; without it pdf.js parses on the main thread */
  pdfWorkerSrc?: string;
  /** overrides the pdf.js import (tests) */
  loadPdfJs?: PdfJsLoader;
  /** overrides DOCX text measurement (tests) */
  measureParagraph?: ParagraphMeasurer;
}

const DEFAULT_MAX_DOWNLOAD = 128 * 1024 * 1024;

function isArrayBufferLike(value: unknown): value is ArrayBuffer | ArrayBufferView {
  return (
    value instanceof ArrayBuffer ||
    (typeof ArrayBuffer !== "undefined" && ArrayBuffer.isView(value as ArrayBufferView))
  );
}

export function isViewerSource(value: unknown): value is ViewerSource {
  if (value === null || value === undefined) return false;
  if (typeof value === "string") return true;
  if (value instanceof URL) return true;
  if (isArrayBufferLike(value)) return true;
  if (typeof Blob !== "undefined" && value instanceof Blob) return true;
  return false;
}

async function bytesFromUrl(url: string, options: LoadOptions): Promise<Uint8Array> {
  const maxBytes = options.maxDownloadBytes ?? DEFAULT_MAX_DOWNLOAD;

  let parsed: URL;
  try {
    parsed = new URL(url, typeof location !== "undefined" ? location.href : "http://localhost/");
  } catch {
    throw new DocumentViewerError("UNSUPPORTED_INPUT", `"${url}" is not a valid URL.`);
  }
  if (!["http:", "https:", "blob:", "data:"].includes(parsed.protocol)) {
    throw new DocumentViewerError(
      "UNSUPPORTED_INPUT",
      `Refusing to load a document from the "${parsed.protocol}" scheme.`,
    );
  }
  if (typeof fetch !== "function") {
    throw new DocumentViewerError("NETWORK_ERROR", "This environment does not provide fetch().");
  }

  let response: Response;
  try {
    response = await fetch(parsed.toString(), {
      credentials: "same-origin",
      ...options.fetchOptions,
      signal: options.signal ?? options.fetchOptions?.signal ?? null,
    });
  } catch (error) {
    if (isAbortError(error)) throw error;
    throw new DocumentViewerError(
      "NETWORK_ERROR",
      "The document could not be downloaded. This is often a network error or a CORS restriction.",
      { cause: error, detail: error instanceof Error ? error.message : String(error) },
    );
  }

  if (!response.ok) {
    throw new DocumentViewerError(
      "NETWORK_ERROR",
      `The document could not be downloaded (HTTP ${response.status} ${response.statusText}).`,
    );
  }

  const declaredLength = Number(response.headers.get("content-length") ?? "");
  if (Number.isFinite(declaredLength) && declaredLength > maxBytes) {
    throw new DocumentViewerError(
      "LIMIT_EXCEEDED",
      `The download is larger than the configured limit of ${Math.round(maxBytes / (1024 * 1024))}MiB.`,
    );
  }

  const buffer = await response.arrayBuffer();
  if (buffer.byteLength > maxBytes) {
    throw new DocumentViewerError(
      "LIMIT_EXCEEDED",
      `The download is larger than the configured limit of ${Math.round(maxBytes / (1024 * 1024))}MiB.`,
    );
  }
  return new Uint8Array(buffer);
}

export async function loadPresentationBytes(
  source: ViewerSource,
  options: LoadOptions = {},
): Promise<Uint8Array> {
  if (source === null || source === undefined) {
    throw new DocumentViewerError("UNSUPPORTED_INPUT", "No document source was provided.");
  }
  if (typeof source === "string") {
    return bytesFromUrl(source, options);
  }
  if (source instanceof URL) {
    return bytesFromUrl(source.href, options);
  }
  if (source instanceof Uint8Array) {
    return source;
  }
  if (source instanceof ArrayBuffer) {
    return new Uint8Array(source).slice();
  }
  if (ArrayBuffer.isView(source)) {
    const view = source as ArrayBufferView;
    return new Uint8Array(view.buffer as ArrayBuffer, view.byteOffset, view.byteLength).slice();
  }
  if (typeof Blob !== "undefined" && source instanceof Blob) {
    const buffer = await source.arrayBuffer();
    return new Uint8Array(buffer);
  }
  throw new DocumentViewerError(
    "UNSUPPORTED_INPUT",
    "The source must be a File, Blob, ArrayBuffer, Uint8Array, URL or URL string.",
  );
}

/** Map parser-level failures onto the public error model. */
export function normalizeError(error: unknown): DocumentViewerError {
  if (error instanceof DocumentViewerError) return error;
  if (isAbortError(error)) {
    return new DocumentViewerError("ABORTED", "Loading was cancelled.");
  }
  if (error instanceof ArchiveError) {
    switch (error.code) {
      case "NOT_A_ZIP":
      case "EMPTY_ARCHIVE":
        return new DocumentViewerError("NOT_A_PRESENTATION", error.message, { cause: error });
      case "ENCRYPTED":
        return new DocumentViewerError("ENCRYPTED", error.message, { cause: error });
      case "ARCHIVE_LIMIT_EXCEEDED":
        return new DocumentViewerError("LIMIT_EXCEEDED", error.message, { cause: error });
      case "CORRUPT_ARCHIVE":
        return new DocumentViewerError("CORRUPT_PACKAGE", error.message, { cause: error, detail: error.detail });
      default:
        return new DocumentViewerError("UNKNOWN", error.message, { cause: error });
    }
  }
  if (error instanceof XmlError) {
    return new DocumentViewerError("MALFORMED_PRESENTATION", "The document contains malformed XML.", {
      cause: error,
      detail: error.message,
    });
  }
  if (error instanceof Error) {
    return new DocumentViewerError("UNKNOWN", error.message || "The document could not be loaded.", {
      cause: error,
    });
  }
  return new DocumentViewerError("UNKNOWN", String(error));
}

/**
 * Parse bytes into a presentation model (`.pptx` only).
 *
 * Parsing itself is synchronous. The caller should yield to the event loop
 * first (the viewer does) so loading state can paint before the work starts.
 * Formats other than `.pptx` go through `parsePackageBytes`/`parseDocumentBytes`
 * in `document.ts`, which decides the format from the bytes.
 */
export function parseBytes(
  bytes: Uint8Array,
  options: Pick<LoadOptions, "limits" | "maxDiagnostics"> = {},
): ParsedPresentation {
  const archive = readArchive(bytes, options.limits ?? {});
  return parsePresentationArchive(archive, {
    limits: options.limits,
    maxDiagnostics: options.maxDiagnostics,
  });
}

/** Convenience: source in, parsed presentation out. */
export async function loadPresentation(
  source: ViewerSource,
  options: LoadOptions = {},
): Promise<ParsedPresentation> {
  const bytes = await loadPresentationBytes(source, options);
  return parseBytes(bytes, options);
}
