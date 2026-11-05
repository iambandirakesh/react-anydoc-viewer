/**
 * Builds a link to a pre-filled GitHub bug report for what the playground is
 * showing. The file itself and its name are never included: the report carries
 * only the format, the counts, the error and the diagnostics the viewer reported,
 * and the person reviews it on GitHub before anything is posted.
 */

import type { DocumentInfo, DocumentViewerErrorType } from "react-anydoc-viewer";

export const GITHUB_URL = "https://github.com/iambandirakesh/react-anydoc-viewer";

/** the bug form's dropdown options (`.github/ISSUE_TEMPLATE/bug_report.yml`) */
const FORMAT_OPTIONS: Record<string, string> = {
  pptx: "PowerPoint (.pptx)",
  docx: "Word (.docx)",
  xlsx: "Excel (.xlsx / .xlsm)",
  pdf: "PDF",
  image: "Image (including TIFF)",
  csv: "CSV / TSV",
  text: "Text (.txt, .json, .md, .html, .log)",
};

/** keeps the URL well under the length GitHub and browsers accept */
const MAX_DIAGNOSTICS = 25;
const MAX_FIELD = 2500;

export interface ReportContext {
  version: string;
  info: DocumentInfo | null;
  error: DocumentViewerErrorType | null;
  /** the playground sample, when one is open */
  sample: string | null;
  /** the extension of the visitor's own file, never its name */
  extension: string | null;
}

function clip(text: string): string {
  return text.length > MAX_FIELD ? `${text.slice(0, MAX_FIELD)}\n… (truncated)` : text;
}

export function issueUrl({ version, info, error, sample, extension }: ReportContext): string {
  const subject = sample ? `the "${sample}" sample` : extension ? `a ${extension} file` : "the demo";
  const environment = [
    `react-anydoc-viewer: ${version} (live demo)`,
    `Browser: ${typeof navigator === "undefined" ? "unknown" : navigator.userAgent}`,
    `Opened: ${subject}`,
    info ? `Format: ${info.format}, ${info.pageCount} page(s)` : null,
  ]
    .filter(Boolean)
    .join("\n");

  const diagnostics = [
    error ? `Error ${error.code}: ${error.message}` : null,
    ...(info?.diagnostics ?? [])
      .slice(0, MAX_DIAGNOSTICS)
      .map(
        (item) =>
          `[${item.severity}] ${item.code}${item.slideIndex === undefined ? "" : ` (page ${item.slideIndex + 1})`}: ${item.message}`,
      ),
    info && info.diagnostics.length > MAX_DIAGNOSTICS ? `… and ${info.diagnostics.length - MAX_DIAGNOSTICS} more` : null,
  ]
    .filter(Boolean)
    .join("\n");

  const params = new URLSearchParams({
    template: "bug_report.yml",
    title: error ? `${error.code} when opening ${subject}` : `Rendering issue: ${subject}`,
    environment: clip(environment),
  });
  if (diagnostics) params.set("diagnostics", clip(diagnostics));
  const format = info ? FORMAT_OPTIONS[info.format] : undefined;
  if (format) params.set("format", format);
  return `${GITHUB_URL}/issues/new?${params.toString()}`;
}
