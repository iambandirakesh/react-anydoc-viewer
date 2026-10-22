/**
 * Public error model.
 *
 * Every failure the viewer can produce is normalised into a {@link DocumentViewerError}
 * with a stable `code`, so applications can branch on the cause without matching
 * on message text.
 */

export type DocumentViewerErrorCode =
  /** the `source` prop was not a supported input type */
  | "UNSUPPORTED_INPUT"
  /** fetching a URL failed (network, CORS, HTTP status) */
  | "NETWORK_ERROR"
  /** the bytes are not a readable document of any supported format */
  | "NOT_A_PRESENTATION"
  /** a real Office/PDF container, but not a format the viewer renders */
  | "UNSUPPORTED_FORMAT"
  /** an encrypted or legacy binary (.ppt) file */
  | "ENCRYPTED"
  /** a configured resource limit was exceeded */
  | "LIMIT_EXCEEDED"
  /** the package is truncated or its directory is unreadable */
  | "CORRUPT_PACKAGE"
  /** the package has no usable slide list */
  | "NO_SLIDES"
  /** the package is missing the parts that make it a presentation */
  | "MALFORMED_PRESENTATION"
  /** the document is unreadable (a damaged PDF or a WordprocessingML package) */
  | "MALFORMED_DOCUMENT"
  /** a slide failed to render */
  | "RENDER_ERROR"
  /** the load was cancelled by a newer request */
  | "ABORTED"
  /** nothing else fits */
  | "UNKNOWN";

export class DocumentViewerError extends Error {
  readonly code: DocumentViewerErrorCode;
  /** technical detail, safe to log but not intended for end users */
  readonly detail?: string;
  readonly cause?: unknown;

  constructor(
    code: DocumentViewerErrorCode,
    message: string,
    options: { detail?: string; cause?: unknown } = {},
  ) {
    super(message);
    this.name = "DocumentViewerError";
    this.code = code;
    this.detail = options.detail;
    this.cause = options.cause;
  }

  toJSON(): { name: string; code: DocumentViewerErrorCode; message: string; detail?: string } {
    const out: { name: string; code: DocumentViewerErrorCode; message: string; detail?: string } = {
      name: this.name,
      code: this.code,
      message: this.message,
    };
    if (this.detail !== undefined) out.detail = this.detail;
    return out;
  }
}

export function isDocumentViewerError(value: unknown): value is DocumentViewerError {
  return value instanceof DocumentViewerError;
}

export type DocumentFormat = "pptx" | "docx" | "pdf" | "image" | "csv" | "text";

/**
 * The formats the viewer can open, for error messages and documentation.
 *
 * This list and the extension table in `src/formats.ts` are kept in step by a
 * unit test and by the package gate, so a format cannot be advertised in one
 * place and missing from the other.
 */
export const SUPPORTED_FORMATS: readonly DocumentFormat[] = [
  "pptx",
  "docx",
  "pdf",
  "image",
  "csv",
  "text",
];

/** Map an {@link AbortError} or an aborted request into a normalised error. */
export function abortedError(): DocumentViewerError {
  return new DocumentViewerError("ABORTED", "Loading was cancelled because a newer presentation was requested.");
}

export function isAbortError(error: unknown): boolean {
  return (
    typeof error === "object" &&
    error !== null &&
    "name" in error &&
    (error as { name?: string }).name === "AbortError"
  );
}

/* ------------------------------------------------------------------ aliases */

/**
 * @deprecated The package was called `react-pptx-viewer` before it also rendered
 * `.docx` and `.pdf`; use {@link DocumentViewerError}. The value is identical, so
 * `instanceof` and `code` checks keep working across both names.
 */
export const PPTViewerError = DocumentViewerError;
/** @deprecated Use {@link DocumentViewerErrorCode}. */
export type PPTViewerError = DocumentViewerError;
/** @deprecated Use {@link DocumentViewerErrorCode}. */
export type PPTViewerErrorCode = DocumentViewerErrorCode;
/** @deprecated Use {@link isDocumentViewerError}. */
export const isPPTViewerError = isDocumentViewerError;
