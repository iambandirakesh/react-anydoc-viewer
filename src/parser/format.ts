/**
 * Document format detection.
 *
 * A `File` may carry a wrong name, a wrong MIME type, or no name at all (an
 * `ArrayBuffer` from `fetch`), so the bytes are the only reliable signal.
 * Detection therefore works the way a file tool does — magic bytes first, then
 * the package's own part names:
 *
 *   1. `%PDF-` anywhere in the leading kilobyte  -> PDF
 *   2. the OLE/CFB signature                     -> legacy or encrypted Office
 *   3. the ZIP signature                         -> an OPC package, whose kind
 *      (presentation, document, spreadsheet) is decided by its part names.
 *
 * Image magic bytes and text classification are handled with this module's
 * output in `document.ts`, because they are not containers: an image is
 * recognised by its header, and a text file by the fact that it decodes as text.
 */

import type { DocumentFormat } from "../errors";
import { supportedExtensionsSentence } from "../formats";

export type ContainerKind =
  /** `%PDF-` header */
  | "pdf"
  /** a ZIP/OPC package */
  | "zip"
  /** an OLE/CFB compound file: legacy `.doc`/`.ppt`/`.xls`, or an encrypted OOXML file */
  | "ole"
  | "unknown";

export interface PackageProbe {
  /** the format the viewer can render, or null when it cannot */
  format: DocumentFormat | null;
  /** short human label for what the package is, used in error messages */
  label: string;
  /** the main part that was located, when known */
  mainPart: string | null;
}

const PDF_HEADER = [0x25, 0x50, 0x44, 0x46, 0x2d]; // %PDF-
const OLE_HEADER = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

function hasBytesAt(bytes: Uint8Array, offset: number, signature: readonly number[]): boolean {
  if (bytes.length < offset + signature.length) return false;
  for (let i = 0; i < signature.length; i += 1) {
    if (bytes[offset + i] !== signature[i]) return false;
  }
  return true;
}

function looksLikeZip(bytes: Uint8Array): boolean {
  if (bytes.length < 4) return false;
  return (
    (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04) ||
    (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x05 && bytes[3] === 0x06) ||
    (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x07 && bytes[3] === 0x08)
  );
}

/**
 * Identify the container. The PDF header is allowed to be preceded by junk
 * (it is defined to appear within the first 1024 bytes), which is what makes
 * detection of PDFs with a byte-order-mark or a stray prefix work.
 */
export function sniffContainer(bytes: Uint8Array): ContainerKind {
  if (bytes.length === 0) return "unknown";
  if (hasBytesAt(bytes, 0, OLE_HEADER)) return "ole";
  if (looksLikeZip(bytes)) return "zip";
  const limit = Math.min(bytes.length - PDF_HEADER.length, 1024);
  for (let offset = 0; offset <= limit; offset += 1) {
    if (hasBytesAt(bytes, offset, PDF_HEADER)) return "pdf";
  }
  return "unknown";
}

const PRESENTATION_MARKERS = ["ppt/presentation.xml"];
const DOCUMENT_MARKERS = ["word/document.xml"];

/**
 * Decide what an OPC package is from the parts it contains.
 *
 * Part names are checked case-insensitively (OPC part names are), and the
 * markers are the main parts of each format, so a renamed `.docx` or a package
 * without a useful file extension is still recognised.
 */
export function probePackage(names: readonly string[]): PackageProbe {
  const lower = new Set(names.map((name) => name.toLowerCase()));
  const has = (marker: string) => lower.has(marker);

  if (PRESENTATION_MARKERS.some(has)) {
    return { format: "pptx", label: "a PowerPoint presentation (.pptx)", mainPart: "ppt/presentation.xml" };
  }
  if (DOCUMENT_MARKERS.some(has)) {
    return { format: "docx", label: "a Word document (.docx)", mainPart: "word/document.xml" };
  }
  if (has("xl/workbook.xml")) {
    return { format: "xlsx", label: "an Excel workbook (.xlsx)", mainPart: "xl/workbook.xml" };
  }
  if (has("xl/workbook.bin")) {
    return { format: null, label: "an Excel binary workbook (.xlsb)", mainPart: "xl/workbook.bin" };
  }
  if (has("mimetype")) {
    return { format: null, label: "an OpenDocument file", mainPart: null };
  }
  return { format: null, label: "an unfamiliar Office package", mainPart: null };
}

/**
 * The message used when a recognised container is not something we can render.
 *
 * The list of formats is generated from the format table (`src/formats.ts`),
 * because a hand-written list is one that eventually disagrees with the code.
 */
export function unsupportedFormatMessage(label: string): string {
  return `This file is ${label}, which the viewer cannot display. Supported formats are ${supportedExtensionsSentence()}.`;
}
