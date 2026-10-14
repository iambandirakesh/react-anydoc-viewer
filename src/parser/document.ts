/**
 * The single document entry point.
 *
 * Every source the viewer accepts ends up here, and the format is decided from
 * the bytes rather than from a file name or MIME type:
 *
 *   %PDF-            -> pdf.js renders pages onto canvases
 *   ZIP/OPC          -> PresentationML (slides) or WordprocessingML (pages)
 *   OLE/CFB          -> legacy or encrypted Office file, refused with a reason
 *   anything else    -> refused with the list of supported formats
 *
 * A package is read once and inspected for its main part, so the same
 * `[Content_Types].xml` / part-name evidence that identifies a presentation also
 * identifies a document — no extension guessing anywhere in the pipeline.
 */

import type { PresentationModel } from "../model/types";
import { DocumentViewerError, type DocumentFormat } from "../errors";
import { readArchive, type ArchiveLimits } from "./archive";
import { probePackage, sniffContainer, unsupportedFormatMessage } from "./format";
import { loadPresentationBytes, normalizeError, type LoadOptions, type ViewerSource } from "./loadPresentation";
import { parsePresentationArchive } from "./presentation";
import { parseDocxArchive, type ParsedDocx } from "./docx";
import { parsePdfDocument, type PdfEngine, type PdfJsLoader } from "./pdf";
import type { MediaStore } from "./media";
import type { PageRenderer } from "../renderer/pageRenderer";
import type { ParagraphMeasurer } from "./docx/model";

export interface DocumentParseOptions {
  limits?: Partial<ArchiveLimits>;
  maxDiagnostics?: number;
  signal?: AbortSignal;
  /** URL of pdf.js's worker module; without it pdf.js parses on the main thread */
  pdfWorkerSrc?: string;
  /** forces a pdf.js build instead of feature-detecting one (tests) */
  pdfEngine?: PdfEngine;
  /** overrides the pdf.js import (tests) */
  loadPdfJs?: PdfJsLoader;
  /** overrides DOCX text measurement (tests) */
  measureParagraph?: ParagraphMeasurer;
}

export interface ParsedDocument {
  format: DocumentFormat;
  model: PresentationModel;
  media: MediaStore;
  /** set by formats whose pages are not a shape list (PDF) */
  renderPage?: PageRenderer;
  dispose(): void;
}

/** What a set of bytes turns out to be, without committing to parsing it. */
export function detectFormat(bytes: Uint8Array): DocumentFormat | "legacy" | "unknown" {
  const container = sniffContainer(bytes);
  if (container === "pdf") return "pdf";
  if (container === "ole") return "legacy";
  if (container !== "zip") return "unknown";
  try {
    const archive = readArchive(bytes);
    return probePackage(archive.names).format ?? "unknown";
  } catch {
    return "unknown";
  }
}

/**
 * Parse a ZIP/OPC package (`.pptx` or `.docx`) synchronously.
 *
 * PDFs need an async engine, so they are rejected here with a message pointing at
 * {@link parseDocumentBytes}; everything else keeps the synchronous fast path the
 * viewer has always used for presentations.
 */
export function parsePackageBytes(bytes: Uint8Array, options: DocumentParseOptions = {}): ParsedDocument {
  const container = sniffContainer(bytes);
  if (container === "pdf") {
    throw new DocumentViewerError("UNSUPPORTED_INPUT", "PDFs must be parsed asynchronously.");
  }
  const archive = readArchive(bytes, options.limits ?? {});
  const probe = probePackage(archive.names);

  if (probe.format === "pptx") {
    const parsed = parsePresentationArchive(archive, {
      limits: options.limits,
      maxDiagnostics: options.maxDiagnostics,
    });
    return { format: "pptx", model: parsed.model, media: parsed.media, dispose: parsed.dispose };
  }
  if (probe.format === "docx") {
    const parsed: ParsedDocx = parseDocxArchive(archive, {
      limits: options.limits,
      maxDiagnostics: options.maxDiagnostics,
      ...(options.measureParagraph ? { measure: options.measureParagraph } : {}),
    });
    return { format: "docx", model: parsed.model, media: parsed.media, dispose: parsed.dispose };
  }
  throw new DocumentViewerError("UNSUPPORTED_FORMAT", unsupportedFormatMessage(probe.label), {
    detail: probe.mainPart ?? undefined,
  });
}

/** Convenience: source in, parsed document out, format decided from the bytes. */
export async function loadDocument(
  source: ViewerSource,
  options: LoadOptions = {},
): Promise<ParsedDocument> {
  try {
    const bytes = await loadPresentationBytes(source, options);
    return await parseDocumentBytes(bytes, options);
  } catch (error) {
    throw normalizeError(error);
  }
}

export async function parseDocumentBytes(
  bytes: Uint8Array,
  options: DocumentParseOptions = {},
): Promise<ParsedDocument> {
  if (sniffContainer(bytes) === "pdf") {
    // `parsePdfDocument` normalises its own failures already.
    const parsed = await parsePdfDocument(bytes, {
      workerSrc: options.pdfWorkerSrc,
      maxDiagnostics: options.maxDiagnostics,
      signal: options.signal,
      ...(options.pdfEngine ? { engine: options.pdfEngine } : {}),
      ...(options.loadPdfJs ? { loadPdfJs: options.loadPdfJs } : {}),
    });
    return {
      format: "pdf",
      model: parsed.model,
      media: parsed.media,
      renderPage: parsed.renderPage,
      dispose: parsed.dispose,
    };
  }
  return parsePackageBytes(bytes, options);
}
