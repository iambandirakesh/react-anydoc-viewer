/**
 * The single source of truth for what the viewer opens.
 *
 * Every user-facing list — error messages, the README table, a file picker's
 * `accept` attribute, a host's server-side allow-list — is derived from
 * {@link FORMAT_DESCRIPTORS} instead of being written out again. Format lists
 * drift when they are copied, and a drifted list is a support ticket.
 *
 * Detection still happens on the bytes (`parser/format.ts`, `parser/document.ts`).
 * Nothing here influences what a file is recognised as; it only describes what
 * the viewer is able to recognise.
 */

import { SUPPORTED_FORMATS, type DocumentFormat } from "./errors";

export interface FormatDescriptor {
  format: DocumentFormat;
  /** short human label, singular */
  label: string;
  /** lower-case extensions, dot included */
  extensions: readonly string[];
  mimeTypes: readonly string[];
}

/**
 * Order matters: this is the order formats are listed to users, from the most
 * common upload to the least.
 */
export const FORMAT_DESCRIPTORS: readonly FormatDescriptor[] = [
  {
    format: "pptx",
    label: "PowerPoint presentation",
    extensions: [".pptx"],
    mimeTypes: ["application/vnd.openxmlformats-officedocument.presentationml.presentation"],
  },
  {
    format: "docx",
    label: "Word document",
    extensions: [".docx"],
    mimeTypes: ["application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
  },
  {
    format: "pdf",
    label: "PDF document",
    extensions: [".pdf"],
    mimeTypes: ["application/pdf"],
  },
  {
    format: "image",
    label: "image",
    // Every one of these is decoded by the image element itself; TIFF, EMF and
    // WMF are detected but deliberately not advertised, because no browser
    // decodes them and the viewer reports them instead of showing a broken box.
    extensions: [".png", ".jpg", ".jpeg", ".gif", ".webp", ".avif", ".bmp", ".svg"],
    mimeTypes: [
      "image/png",
      "image/jpeg",
      "image/gif",
      "image/webp",
      "image/avif",
      "image/bmp",
      "image/svg+xml",
    ],
  },
  {
    format: "csv",
    label: "delimited text table",
    // The delimiter is read from the content, so a tab-separated file needs no
    // separate format.
    extensions: [".csv", ".tsv"],
    mimeTypes: ["text/csv", "text/tab-separated-values"],
  },
  {
    format: "text",
    label: "text file",
    // Everything here is one document kind: bytes that decode as text. JSON is
    // pretty-printed, everything else — HTML included — is paginated as monospace
    // text and never interpreted. Markdown is listed here because that is what
    // it is today — read as text, with a diagnostic saying so — so a file picker
    // generated from this table offers it. It moves to its own descriptor when
    // Markdown is rendered as a document rather than shown as source.
    extensions: [".txt", ".log", ".json", ".html", ".htm", ".md", ".markdown"],
    mimeTypes: ["text/plain", "text/html", "application/json", "text/markdown"],
  },
];

/** The descriptor for a format the viewer advertises. */
export function descriptorFor(format: DocumentFormat): FormatDescriptor | null {
  return FORMAT_DESCRIPTORS.find((descriptor) => descriptor.format === format) ?? null;
}

/** Every extension the viewer can open, lower-case and dot-prefixed. */
export const SUPPORTED_EXTENSIONS: readonly string[] = FORMAT_DESCRIPTORS.flatMap(
  (descriptor) => descriptor.extensions,
);

/**
 * A ready-made value for `<input type="file" accept="…">`, so a picker cannot
 * offer a file the viewer will refuse (or hide one it could have opened).
 */
export const ACCEPT_ATTRIBUTE: string = FORMAT_DESCRIPTORS.flatMap((descriptor) => [
  ...descriptor.extensions,
  ...descriptor.mimeTypes,
]).join(",");

/**
 * An English list of the supported formats ("`.pptx`, `.docx` and `.pdf`"), built
 * from the same table so a new format cannot be forgotten in an error message.
 */
export function supportedFormatsSentence(format: (descriptor: FormatDescriptor) => string): string {
  const items = FORMAT_DESCRIPTORS.map(format);
  if (items.length === 0) return "";
  if (items.length === 1) return items[0] as string;
  const head = items.slice(0, -1).join(", ");
  return `${head} and ${items[items.length - 1] as string}`;
}

/** Extensions of the formats the viewer opens, for error messages. */
export function supportedExtensionsSentence(): string {
  return supportedFormatsSentence((descriptor) => descriptor.extensions.join("/"));
}

/** Assert at build time that every advertised format has a descriptor. */
export function missingDescriptors(): DocumentFormat[] {
  return SUPPORTED_FORMATS.filter((format) => descriptorFor(format) === null);
}
