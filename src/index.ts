/**
 * react-anydoc-viewer
 *
 * A free, MIT-licensed React component that renders PowerPoint (.pptx), Word
 * (.docx) and PDF files in the browser. Parsing and rendering happen entirely
 * client-side; there are no paid APIs, external services, watermarks or branding
 * overlays. The format is detected from the bytes, so one component serves all
 * three.
 *
 * ```tsx
 * import { DocumentViewer } from "react-anydoc-viewer";
 * import "react-anydoc-viewer/styles.css";
 *
 * <DocumentViewer file="/deck.pptx" showControls showThumbnails fit="width" />
 * <DocumentViewer file="/report.docx" />
 * <DocumentViewer file="/paper.pdf" pdfWorkerSrc={workerUrl} />
 * ```
 *
 * The public surface is deliberately small: one component, its props, navigation
 * handle, diagnostics and error model. Parser internals are not exported, so
 * they can change without a breaking release.
 */

import { DocumentViewer } from "./components/DocumentViewer";

export { DocumentViewer };
export default DocumentViewer;

// The names from the package's previous life as `react-pptx-viewer`. They are
// aliases, not copies: the same component, class and types, so a consumer can
// migrate file by file without two viewer implementations in the tree.
export { PPTViewer } from "./components/DocumentViewer";
export {
  DocumentViewerError,
  isDocumentViewerError,
  /** @deprecated Use {@link DocumentViewerError}. */
  PPTViewerError,
  /** @deprecated Use {@link isDocumentViewerError}. */
  isPPTViewerError,
} from "./errors";
export { DEFAULT_LABELS } from "./components/labels";
export { DEFAULT_ARCHIVE_LIMITS } from "./parser/archive";
export { supportedPresetNames } from "./parser/geometry";
/** Which pdf.js build this environment needs; pick `pdfWorkerSrc` to match it. */
export { detectPdfEngine } from "./parser/pdf";

export type { PdfEngine } from "./parser/pdf";
export type { ArchiveLimits } from "./parser/archive";
export type {
  DocumentFormat,
  DocumentViewerErrorCode,
  /** @deprecated Use {@link DocumentViewerErrorCode}. */
  PPTViewerErrorCode,
} from "./errors";
export { SUPPORTED_FORMATS } from "./errors";
export type {
  FitMode,
  DocumentDiagnostic,
  DocumentSource,
  DocumentViewerErrorType,
  DocumentViewerHandle,
  DocumentViewerProps,
  DocumentInfo,
  ViewerLabels,
  /** @deprecated Use {@link DocumentViewerProps}. */
  PPTViewerProps,
  /** @deprecated Use {@link DocumentViewerHandle}. */
  PPTViewerHandle,
  /** @deprecated Use {@link DocumentViewerErrorType}. */
  PPTViewerErrorType,
  /** @deprecated Use {@link DocumentSource}. */
  PptxSource,
  /** @deprecated Use {@link DocumentDiagnostic}. */
  PptxDiagnostic,
  /** @deprecated Use {@link DocumentInfo}. */
  PresentationInfo,
} from "./types/public";

/** Package version, useful for diagnostics and bug reports. */
export const version = "0.1.0" as const;
