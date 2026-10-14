/**
 * PDF adapter.
 *
 * A PDF is not a shape list, so it is not parsed into the model: pdf.js renders
 * each page to a canvas and the viewer's page-renderer seam draws it. That keeps
 * the format honest — no half-reimplementation of a PDF engine — and reuses the
 * viewer's navigation, thumbnails, zoom and presentation mode unchanged.
 *
 * Loading is lazy in every sense:
 *  - `pdfjs-dist` is imported only when a PDF is actually opened, so `.pptx` and
 *    `.docx` users never download it;
 *  - pages are rendered when they become visible, not when the file is loaded;
 *  - the module is never imported during server rendering.
 *
 * pdf.js normally runs its parser in a Web Worker. A published library cannot
 * conjure a worker URL out of a consumer's bundler, so by default pdf.js is told
 * to use its main-thread fallback (its own documented "fake worker"), which works
 * wherever the module resolves at all. `pdfWorkerSrc` switches to a real worker,
 * which is what a production app should do.
 *
 * pdf.js ships two builds: the modern one and a `legacy` build that bundles the
 * polyfills for older engines (core-js). The modern build assumes `Float16Array`,
 * `Uint8Array.prototype.toHex`, `Promise.try` and friends, so on an older
 * Chromium — an Electron shell, for example — it fails at runtime with
 * "a.toHex is not a function". The build is therefore chosen by feature
 * detection rather than by browser version sniffing; see {@link detectPdfEngine}.
 */

import type { PresentationModel, SlideModel } from "../model/types";
import { DocumentViewerError } from "../errors";
import type { ArchiveLimits } from "./archive";
import { Diagnostics, DEFAULT_MAX_DIAGNOSTICS } from "./diagnostics";
import { createEmptyMediaStore, type MediaStore } from "./media";
import { DEFAULT_COLOR_MAP, createEmptyTheme } from "./theme";
import { EMU_PER_POINT } from "./units";
import type { PageRenderer } from "../renderer/pageRenderer";
import { createPdfPageRenderer, type PdfDocumentLike } from "../components/PdfPage";

interface PdfViewportLike {
  width: number;
  height: number;
}

interface PdfPageProxyLike {
  getViewport(params: { scale: number }): PdfViewportLike;
}

interface PdfDocumentProxyLike {
  numPages: number;
  getPage(pageNumber: number): Promise<PdfPageProxyLike>;
  getMetadata(): Promise<{ info?: { Title?: string; Author?: string } }>;
  destroy(): Promise<void>;
}

interface PdfJsLike {
  GlobalWorkerOptions: { workerSrc: string };
  getDocument(params: Record<string, unknown>): { promise: Promise<unknown> };
}

export type PdfJsLoader = (engine: PdfEngine) => Promise<PdfJsLike>;

/** Which pdf.js build the environment needs. */
export type PdfEngine = "modern" | "legacy";

export interface PdfParseOptions {
  /** URL of `pdf.worker.min.mjs`; without it pdf.js parses on the main thread */
  workerSrc?: string;
  /** forces a build instead of feature-detecting one (tests) */
  engine?: PdfEngine;
  limits?: Partial<ArchiveLimits>;
  maxDiagnostics?: number;
  signal?: AbortSignal;
  /** overrides the module import (tests, alternative builds) */
  loadPdfJs?: PdfJsLoader;
}

export interface ParsedPdf {
  model: PresentationModel;
  media: MediaStore;
  renderPage: PageRenderer;
  dispose(): void;
}

/** Pages whose size is compared before the first page's geometry is assumed. */
const SIZE_SAMPLE_LIMIT = 200;

const defaultLoader: PdfJsLoader = (engine) =>
  (engine === "legacy"
    ? import("pdfjs-dist/legacy/build/pdf.mjs")
    : import("pdfjs-dist")) as unknown as Promise<PdfJsLike>;

/** The main-thread ("fake worker") module matching the chosen build. */
function loadFakeWorker(engine: PdfEngine): Promise<unknown> {
  return engine === "legacy"
    ? import("pdfjs-dist/legacy/build/pdf.worker.min.mjs")
    : import("pdfjs-dist/build/pdf.worker.min.mjs");
}

type GlobalWithPdfWorker = typeof globalThis & { pdfjsWorker?: unknown };
type ModernApi = { toHex?: unknown };
type PromiseExtras = { try?: unknown; withResolvers?: unknown };
type ArrayExtras = { at?: unknown; findLast?: unknown };

/**
 * The primitives the modern pdf.js build uses. They are newer than the rest of
 * the library's target, so their absence — not a user-agent string — decides
 * which build runs. The legacy build bundles polyfills for exactly these.
 */
const MODERN_BASELINE: ReadonlyArray<readonly [string, () => boolean]> = [
  ["Float16Array", () => typeof (globalThis as { Float16Array?: unknown }).Float16Array === "function"],
  ["Uint8Array.prototype.toHex", () => typeof (Uint8Array.prototype as ModernApi).toHex === "function"],
  ["Promise.withResolvers", () => typeof (Promise as PromiseExtras).withResolvers === "function"],
  ["Promise.try", () => typeof (Promise as PromiseExtras).try === "function"],
  ["Array.prototype.at", () => typeof (Array.prototype as ArrayExtras).at === "function"],
  ["Array.prototype.findLast", () => typeof (Array.prototype as ArrayExtras).findLast === "function"],
  ["URL.parse", () => typeof (URL as { parse?: unknown }).parse === "function"],
  ["structuredClone", () => typeof globalThis.structuredClone === "function"],
];

export interface PdfEngineChoice {
  engine: PdfEngine;
  /** the modern-build features this environment lacks, in detection order */
  missing: string[];
}

/**
 * The answer is frozen on first use, and the first use always precedes the
 * `pdfjs-dist` import. The legacy build's own polyfills patch the prototypes this
 * detection looks at, so re-measuring afterwards would report a modern
 * environment and pick the wrong build for the rest of the page's life.
 */
let cachedChoice: PdfEngineChoice | null = null;

export function pdfEngineChoice(): PdfEngineChoice {
  if (cachedChoice) return cachedChoice;
  const missing: string[] = [];
  for (const [name, supported] of MODERN_BASELINE) {
    let present = false;
    try {
      present = supported();
    } catch {
      present = false;
    }
    if (!present) missing.push(name);
  }
  cachedChoice = { engine: missing.length === 0 ? "modern" : "legacy", missing };
  return cachedChoice;
}

/** Names of the modern-build features this environment does not provide. */
export function missingModernPdfFeatures(): string[] {
  return [...pdfEngineChoice().missing];
}

/**
 * Which pdf.js build this environment needs. Exported so an application can pick
 * the matching worker file for `pdfWorkerSrc` — the builds are not
 * interchangeable: pdf.js refuses to pair an API with a worker from the other
 * build, so a modern worker URL would fail wherever the legacy build is used.
 */
export function detectPdfEngine(): PdfEngine {
  return pdfEngineChoice().engine;
}

async function loadPdfJs(
  options: PdfParseOptions,
): Promise<{ pdfjs: PdfJsLike; engine: PdfEngine; mainThread: boolean }> {
  const engine = options.engine ?? pdfEngineChoice().engine;
  const loader = options.loadPdfJs ?? defaultLoader;
  let pdfjs: PdfJsLike;
  try {
    pdfjs = await loader(engine);
  } catch (error) {
    throw new DocumentViewerError(
      "UNKNOWN",
      "PDF support could not be loaded. The 'pdfjs-dist' package must be installed and resolvable by your bundler.",
      { cause: error, detail: error instanceof Error ? error.message : String(error) },
    );
  }

  const scope = globalThis as GlobalWithPdfWorker;
  if (options.workerSrc) {
    // An explicit worker always wins, including after an earlier main-thread load.
    delete scope.pdfjsWorker;
    pdfjs.GlobalWorkerOptions.workerSrc = options.workerSrc;
    return { pdfjs, engine, mainThread: false };
  }

  if (!scope.pdfjsWorker) {
    try {
      // pdf.js looks for this global and, when it finds it, parses in-process.
      scope.pdfjsWorker = await loadFakeWorker(engine);
    } catch {
      // Let pdf.js try its own worker resolution.
    }
  }
  return { pdfjs, engine, mainThread: Boolean(scope.pdfjsWorker) };
}

export function normalizePdfError(error: unknown): DocumentViewerError {
  const name = (error as { name?: string })?.name ?? "";
  const message = error instanceof Error ? error.message : String(error);
  if (name === "PasswordException") {
    return new DocumentViewerError("ENCRYPTED", "This PDF is password-protected and cannot be opened.", { detail: message });
  }
  if (name === "InvalidPDFException" || name === "MissingPDFException" || name === "UnexpectedResponseException") {
    return new DocumentViewerError("MALFORMED_DOCUMENT", "The file is not a readable PDF.", { detail: message });
  }
  if (name === "AbortError" || name === "RenderingCancelledException") {
    return new DocumentViewerError("ABORTED", "Loading was cancelled.");
  }
  return new DocumentViewerError("UNKNOWN", "The PDF could not be opened.", { cause: error, detail: message });
}

export async function parsePdfDocument(bytes: Uint8Array, options: PdfParseOptions = {}): Promise<ParsedPdf> {
  const diagnostics = new Diagnostics(options.maxDiagnostics ?? DEFAULT_MAX_DIAGNOSTICS);
  const { pdfjs, engine, mainThread } = await loadPdfJs(options);
  if (engine === "legacy") {
    diagnostics.add({
      severity: "info",
      code: "pdf.engine.legacy",
      message: `This environment is missing ${missingModernPdfFeatures().join(
        ", ",
      )}, so pdf.js's legacy build is used. Publish pdfjs-dist/legacy/build/pdf.worker.min.mjs when passing pdfWorkerSrc.`,
    });
  }
  if (mainThread) {
    diagnostics.add({
      severity: "info",
      code: "pdf.worker.main-thread",
      message:
        "Pages are rendered on the main thread. Pass pdfWorkerSrc to use a Web Worker and keep the interface responsive.",
    });
  }

  /** One attempt at opening the file; a copy of the bytes is handed to pdf.js. */
  const open = () =>
    pdfjs
      .getDocument({
        // pdf.js transfers (and detaches) the buffer to its worker, which would
        // otherwise take the caller's bytes with it.
        data: bytes.slice(),
        isEvalSupported: false,
        disableFontFace: false,
        signal: options.signal ?? null,
      })
      .promise as Promise<unknown>;

  let document: PdfDocumentProxyLike;
  try {
    document = (await open()) as PdfDocumentProxyLike;
  } catch (error) {
    const failure = normalizePdfError(error);
    // A worker the viewer cannot use — the other build, a URL that 404s, a
    // blocked cross-origin script — must not cost the file. Retry in-process and
    // report the downgrade, the same way a missing worker is reported.
    if (!options.workerSrc || failure.code !== "UNKNOWN") throw failure;
    diagnostics.add({
      severity: "warning",
      code: "pdf.engine.worker",
      message: `The configured pdf.js worker could not be used (${
        failure.detail ?? failure.message
      }), so pages are rendered without a worker. Publish the worker matching this environment — see detectPdfEngine().`,
    });
    pdfjs.GlobalWorkerOptions.workerSrc = "";
    const scope = globalThis as GlobalWithPdfWorker;
    try {
      scope.pdfjsWorker = await loadFakeWorker(engine);
    } catch {
      // Let pdf.js try its own resolution for the retry as well.
    }
    try {
      document = (await open()) as PdfDocumentProxyLike;
    } catch (retryError) {
      throw normalizePdfError(retryError);
    }
  }

  const pageCount = document.numPages;
  if (pageCount <= 0) {
    await document.destroy().catch(() => undefined);
    throw new DocumentViewerError("MALFORMED_DOCUMENT", "The PDF has no pages.");
  }

  let geometry = { widthEmu: 0, heightEmu: 0 };
  try {
    const first = await document.getPage(1);
    const viewport = first.getViewport({ scale: 1 });
    geometry = {
      widthEmu: Math.round(viewport.width * EMU_PER_POINT),
      heightEmu: Math.round(viewport.height * EMU_PER_POINT),
    };
    const sample = Math.min(pageCount, SIZE_SAMPLE_LIMIT);
    for (let pageNumber = 2; pageNumber <= sample; pageNumber += 1) {
      const page = await document.getPage(pageNumber);
      const size = page.getViewport({ scale: 1 });
      if (Math.abs(size.width - viewport.width) > 0.5 || Math.abs(size.height - viewport.height) > 0.5) {
        diagnostics.add({
          severity: "warning",
          code: "pdf.page-size",
          message: `Page ${pageNumber} has a different size; all pages are shown at the first page's size.`,
          slideIndex: pageNumber - 1,
        });
        break;
      }
    }
  } catch (error) {
    await document.destroy().catch(() => undefined);
    throw normalizePdfError(error);
  }

  let title: string | null = null;
  let author: string | null = null;
  try {
    const metadata = await document.getMetadata();
    const info = metadata?.info;
    title = info?.Title && info.Title.trim().length > 0 ? info.Title.trim() : null;
    author = info?.Author && info.Author.trim().length > 0 ? info.Author.trim() : null;
  } catch {
    // Metadata is optional; a damaged info dictionary is not an error.
  }

  const slides: SlideModel[] = [];
  for (let index = 0; index < pageCount; index += 1) {
    slides.push({
      index,
      partName: `pdf#page${index + 1}`,
      name: null,
      slideshowHidden: false,
      // The page's content is drawn by the page renderer, not from shapes.
      shapes: [],
      inheritedShapes: [],
      background: { type: "solid", color: { r: 255, g: 255, b: 255, a: 1 } },
      diagnostics: [],
    });
  }

  const model: PresentationModel = {
    width: geometry.widthEmu,
    height: geometry.heightEmu,
    slides,
    theme: createEmptyTheme(),
    colorMap: DEFAULT_COLOR_MAP,
    textStyles: { title: [], body: [], other: [] },
    title,
    author,
    partCount: pageCount,
    diagnostics: diagnostics.list(),
  };

  const media = createEmptyMediaStore();
  // pdf.js's document proxy satisfies the renderer's structural contract.
  const renderPage = createPdfPageRenderer(document as unknown as PdfDocumentLike);

  return {
    model,
    media,
    renderPage,
    dispose: () => {
      media.dispose();
      void document.destroy().catch(() => undefined);
    },
  };
}
