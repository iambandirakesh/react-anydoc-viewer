/**
 * The public API surface.
 *
 * Nothing here exposes parser internals: a consumer gets props, a navigation
 * handle, load metadata, diagnostics and the error model, and nothing else.
 */

import type { CSSProperties, ReactNode, Ref } from "react";
import type { Diagnostic } from "../model/types";
import type { ArchiveLimits } from "../parser/archive";
import type { ViewerSource } from "../parser/loadPresentation";
import type { DocumentFormat, DocumentViewerError } from "../errors";

/** Accepted document inputs: File, Blob, ArrayBuffer, Uint8Array, URL or URL string. */
export type DocumentSource = ViewerSource;

/** Diagnostics reported for a document, used by `onDiagnostics` and the debug panel. */
export type DocumentDiagnostic = Diagnostic;

export type DocumentViewerErrorType = DocumentViewerError;

/**
 * How the slide is scaled inside its container.
 * - `width` — fill the container width
 * - `height` — fill the container height
 * - `page` — fit the whole slide inside the container (default)
 * - `actual` — 100%
 * - `none` — no automatic scaling; the container scrolls
 */
export type FitMode = "width" | "height" | "page" | "actual" | "none";

export interface ViewerLabels {
  previous: string;
  next: string;
  first: string;
  last: string;
  zoomIn: string;
  zoomOut: string;
  fitWidth: string;
  fitHeight: string;
  fitPage: string;
  actualSize: string;
  /** label of the zoom percentage dropdown, which also holds the fit modes */
  zoomMenu: string;
  fullscreen: string;
  exitFullscreen: string;
  present: string;
  exitPresentation: string;
  /** shown briefly when presentation mode starts */
  presentationHint: string;
  toggleThumbnails: string;
  thumbnails: string;
  slideOf: (current: number, total: number) => string;
  loading: string;
  errorTitle: string;
  retry: string;
  diagnostics: string;
  diagnosticsClose: string;
  diagnosticsEmpty: string;
  noMeasure: string;
  emptySlide: string;
  renderError: string;
  viewerRegion: string;
}

/** Metadata handed to `onLoad`. */
export interface DocumentInfo {
  /** the format that was detected from the bytes */
  format: DocumentFormat;
  /** number of pages: slides for .pptx, pages for .docx/.xlsx/.pdf/images/text */
  slideCount: number;
  /** slide width in EMU (914400 EMU = 1 inch) */
  widthEmu: number;
  /** slide height in EMU */
  heightEmu: number;
  /** slide width / height */
  aspectRatio: number;
  /** document title from the package properties, when present */
  title: string | null;
  /** diagnostics collected while parsing */
  diagnostics: DocumentDiagnostic[];
}

export interface DocumentViewerHandle {
  /** go to a 0-based slide index (clamped) */
  goToSlide: (index: number) => void;
  next: () => void;
  previous: () => void;
  first: () => void;
  last: () => void;
  getCurrentSlide: () => number;
  getSlideCount: () => number;
  setZoom: (scale: number) => void;
  zoomIn: () => void;
  zoomOut: () => void;
  /** return to the fit mode given by the `fit` prop */
  resetZoom: () => void;
  /** current effective scale (1 === 100%) */
  getZoom: () => number;
  enterFullscreen: () => void;
  exitFullscreen: () => void;
  toggleFullscreen: () => void;
  /** start the chrome-free slideshow, from the current slide */
  enterPresentation: () => void;
  exitPresentation: () => void;
  togglePresentation: () => void;
  isPresenting: () => boolean;
  /** re-download and re-parse the current source */
  reload: () => void;
}

export interface DocumentViewerProps {
  /** the document to display: .pptx, .docx, .xlsx, .pdf, an image, a table or text */
  source?: DocumentSource | null;
  /** alias of `source`, matching the common `file="..."` usage */
  file?: DocumentSource | null;
  /** controlled current slide (0-based) */
  slide?: number;
  /** uncontrolled initial slide */
  defaultSlide?: number;
  /** show the toolbar (Previous/Next/zoom/fullscreen) */
  showControls?: boolean;
  /** alias of `showControls` */
  showToolbar?: boolean;
  /** show the thumbnail sidebar initially */
  showThumbnails?: boolean;
  /** allow the sidebar to be toggled at runtime (needs `showControls`) */
  allowThumbnailToggle?: boolean;
  /** controlled presentation (slideshow) mode */
  presentation?: boolean;
  /** uncontrolled initial presentation mode */
  defaultPresentation?: boolean;
  /** show the Present control in the toolbar (default `true`) */
  showPresentationControl?: boolean;
  fit?: FitMode;
  /** controlled zoom scale; `1` is 100% */
  zoom?: number;
  /** uncontrolled initial zoom scale */
  defaultZoom?: number;
  /** enable Left/Right/Home/End/PageUp/PageDown/Space/+/−/f shortcuts */
  keyboardNavigation?: boolean;
  /** light or dark viewer chrome */
  theme?: "light" | "dark";
  /** convenient single switch for a dark theme */
  dark?: boolean;
  className?: string;
  style?: CSSProperties;
  /** message used when a page genuinely has no content */
  emptySlideMessage?: string;
  labels?: Partial<ViewerLabels>;
  /** show the diagnostics panel and collect every warning */
  debug?: boolean;
  /** archive safety limits */
  limits?: Partial<ArchiveLimits>;
  /** passed through to `fetch` for URL sources */
  fetchOptions?: RequestInit;
  /** maximum bytes accepted from a URL response */
  maxDownloadBytes?: number;
  /**
   * How to read a text file whose content is ambiguous.
   *
   * A `.txt`, `.csv`, `.md` and `.json` file are all "bytes that decode as text",
   * so the viewer decides from the content. Pass `"csv"`, `"markdown"`,
   * `"json"`, `"html"` or `"text"` when the host already knows — for example
   * from the upload's declared type — and the guess is not made. The bytes still
   * decide whether the file can be read at all.
   */
  contentHint?: "text" | "csv" | "json" | "html" | "markdown";
  /**
   * URL of pdf.js's worker module (`pdf.worker.min.mjs`).
   *
   * Without it pdf.js parses PDFs on the main thread, which needs no bundler
   * configuration but can stutter while a page is prepared. Publishing the worker
   * and passing its URL is the recommended production setup. pdf.js ships two
   * builds and refuses to pair an API with the other one's worker, so publish the
   * worker that `detectPdfEngine()` reports for the target environment.
   */
  pdfWorkerSrc?: string;
  /** called after a document parses successfully */
  onLoad?: (info: DocumentInfo) => void;
  /** called on any load or parse failure */
  onError?: (error: DocumentViewerErrorType) => void;
  /** called whenever the visible slide changes */
  onSlideChange?: (index: number) => void;
  /** called with the presentation diagnostics after parsing */
  onDiagnostics?: (diagnostics: DocumentDiagnostic[]) => void;
  /** called when the zoom scale changes */
  onZoomChange?: (scale: number) => void;
  /** called when the fit mode changes through the toolbar */
  onFitChange?: (fit: FitMode) => void;
  /** called when presentation mode is entered or left */
  onPresentationChange?: (active: boolean) => void;
  /** replaces the built-in per-slide error panel */
  renderSlideError?: (error: Error, slideIndex: number) => ReactNode;
  /** replaces the built-in loading indicator */
  renderLoader?: () => ReactNode;
  /** replaces the built-in error state */
  renderError?: (error: DocumentViewerErrorType, retry: () => void) => ReactNode;
  ref?: Ref<DocumentViewerHandle>;
}

/* ------------------------------------------------------------------ aliases */

/** @deprecated Use {@link DocumentViewerProps}. */
export type PPTViewerProps = DocumentViewerProps;
/** @deprecated Use {@link DocumentViewerHandle}. */
export type PPTViewerHandle = DocumentViewerHandle;
/** @deprecated Use {@link DocumentViewerErrorType}. */
export type PPTViewerErrorType = DocumentViewerErrorType;
/** @deprecated Use {@link DocumentSource}. */
export type PptxSource = DocumentSource;
/** @deprecated Use {@link DocumentDiagnostic}. */
export type PptxDiagnostic = DocumentDiagnostic;
/** @deprecated Use {@link DocumentInfo}. */
export type PresentationInfo = DocumentInfo;
