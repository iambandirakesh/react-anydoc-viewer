/**
 * OPC package (ZIP) reading with hard resource limits.
 *
 * Constraints enforced here:
 *  - input size, entry count, per-entry and total uncompressed size, and a
 *    compression-ratio heuristic (guards against decompression bombs);
 *  - entry names are normalised and validated, then looked up by lower-cased
 *    name because OPC part-name matching in Office is case-insensitive;
 *  - only the parts the renderer actually needs are ever inflated, so a deck
 *    full of embedded spreadsheets does not cost memory.
 */

import { unzipSync, type UnzipFileInfo } from "fflate";
import { XmlError } from "./xml";
import { OLE_LABELS, sniffOleKind } from "./ole";

export interface ArchiveLimits {
  /** maximum size of the compressed input bytes */
  maxCompressedBytes: number;
  maxEntries: number;
  maxEntryBytes: number;
  maxTotalBytes: number;
  /** rejected when total uncompressed / compressed exceeds this */
  maxCompressionRatio: number;
  /** maximum size of a single XML part after decoding */
  maxXmlBytes: number;
}

export const DEFAULT_ARCHIVE_LIMITS: ArchiveLimits = {
  maxCompressedBytes: 128 * 1024 * 1024,
  maxEntries: 8000,
  maxEntryBytes: 128 * 1024 * 1024,
  maxTotalBytes: 512 * 1024 * 1024,
  maxCompressionRatio: 200,
  maxXmlBytes: 48 * 1024 * 1024,
};

export type ArchiveErrorCode =
  | "NOT_A_ZIP"
  | "ENCRYPTED"
  /** a pre-2007 binary Office file, which is identified but not rendered */
  | "LEGACY_BINARY"
  | "ARCHIVE_LIMIT_EXCEEDED"
  | "CORRUPT_ARCHIVE"
  | "EMPTY_ARCHIVE";

export class ArchiveError extends Error {
  constructor(
    readonly code: ArchiveErrorCode,
    message: string,
    readonly detail?: string,
  ) {
    super(message);
    this.name = "ArchiveError";
  }
}

export interface ArchiveEntry {
  /** normalised POSIX-style path as stored in the archive */
  name: string;
  bytes: Uint8Array;
}

export interface Archive {
  /** lower-cased part name -> entry */
  entries: Map<string, ArchiveEntry>;
  /** all names present in the package (not necessarily inflated) */
  names: string[];
  limits: ArchiveLimits;
}

export function isSafeEntryName(name: string): boolean {
  if (name.length === 0 || name.length > 512) return false;
  if (name.startsWith("/") || name.startsWith("\\")) return false;
  if (name.includes("\\")) return false;
  if (name.includes("..")) return false;
  // Control characters are never valid in part names and are a common smuggling
  // trick for logs and diagnostics.
  for (let i = 0; i < name.length; i += 1) {
    const code = name.charCodeAt(i);
    if (code < 0x20 || code === 0x7f) return false;
  }
  return true;
}

function normalizeName(raw: string): string {
  let name = raw;
  while (name.startsWith("./")) name = name.slice(2);
  return name.replace(/\/{2,}/g, "/");
}

/** OLE/CFB container: what an encrypted Office file looks like. */
function looksLikeOleContainer(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0 &&
    bytes[4] === 0xa1 &&
    bytes[5] === 0xb1 &&
    bytes[6] === 0x1a &&
    bytes[7] === 0xe1
  );
}
function looksLikeZip(bytes: Uint8Array): boolean {
  if (bytes.length < 4) return false;
  return (
    (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04) ||
    (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x05 && bytes[3] === 0x06) ||
    (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x07 && bytes[3] === 0x08)
  );
}

export type ArchiveSelector = (name: string) => boolean;

/** Parts the viewer needs. Everything else is never inflated. */
export const USEFUL_PART: RegExp[] = [
  /^\[content_types\]\.xml$/,
  /\.rels$/,
  // PresentationML
  /^ppt\/(presentation|presprops|tablestyles|viewprops)\.xml$/,
  /^ppt\/(slides|slidelayouts|slidemasters|theme|charts|media|tags|ink)\//,
  // WordprocessingML
  /^word\/(document|styles|numbering)\d*\.xml$/,
  /^word\/media\//,
  /^docprops\/core\.xml$/,
];

const EXCLUDED_PART: RegExp[] = [
  /^ppt\/notesslides\//,
  /^ppt\/notesmasters\//,
  /^ppt\/embeddings\//,
  /^ppt\/diagrams\//,
  /^ppt\/comments/i,
  /^ppt\/people\.xml$/,
  /^ppt\/handoutmasters\//,
];

export function isUsefulPart(name: string): boolean {
  const lower = name.toLowerCase();
  if (EXCLUDED_PART.some((re) => re.test(lower))) return false;
  return USEFUL_PART.some((re) => re.test(lower));
}

interface DeclaredEntry {
  name: string;
  originalSize: number;
  compressedSize: number;
  compression: number;
}

function preflight(bytes: Uint8Array, limits: ArchiveLimits): DeclaredEntry[] {
  const declared: DeclaredEntry[] = [];
  try {
    unzipSync(bytes, {
      filter: (file: UnzipFileInfo) => {
        declared.push({
          name: file.name,
          originalSize: Number(file.originalSize ?? 0),
          compressedSize: Number(file.size ?? 0),
          compression: Number(file.compression ?? 0),
        });
        return false;
      },
    });
  } catch (error) {
    throw new ArchiveError(
      "CORRUPT_ARCHIVE",
      "The archive directory could not be read.",
      error instanceof Error ? error.message : String(error),
    );
  }
  return declared;
}

export function readArchive(bytes: Uint8Array, overrides: Partial<ArchiveLimits> = {}): Archive {
  const limits: ArchiveLimits = { ...DEFAULT_ARCHIVE_LIMITS, ...overrides };

  if (bytes.length === 0) {
    throw new ArchiveError("EMPTY_ARCHIVE", "The file is empty.");
  }
  if (looksLikeOleContainer(bytes)) {
    // A legacy `.doc`/`.ppt`/`.xls` and an encrypted OOXML package share this
    // container, and they are different problems for the user, so the stream
    // names decide which one is reported.
    const kind = sniffOleKind(bytes);
    if (kind === "excel" || kind === "word" || kind === "powerpoint") {
      throw new ArchiveError(
        "LEGACY_BINARY",
        `This file is ${OLE_LABELS[kind]}, which is written in the pre-2007 binary format the viewer cannot display. Open an exported .xlsx, .docx or .pptx instead.`,
      );
    }
    if (kind === "encrypted") {
      throw new ArchiveError(
        "ENCRYPTED",
        "This document is encrypted or password-protected, so its contents cannot be read.",
      );
    }
    throw new ArchiveError(
      "ENCRYPTED",
      "This looks like an encrypted or legacy binary (.doc, .ppt, .xls) Office file. Encrypted documents and the pre-2007 binary formats are not supported.",
    );
  }
  if (!looksLikeZip(bytes)) {
    throw new ArchiveError(
      "NOT_A_ZIP",
      "The file is neither a ZIP-based Office package (.pptx, .docx) nor a PDF; the PDF signature was not found either.",
    );
  }
  if (bytes.length > limits.maxCompressedBytes) {
    throw new ArchiveError(
      "ARCHIVE_LIMIT_EXCEEDED",
      `The file is larger than the configured limit of ${Math.round(limits.maxCompressedBytes / (1024 * 1024))}MiB.`,
    );
  }

  const declared = preflight(bytes, limits);
  if (declared.length === 0) {
    throw new ArchiveError("EMPTY_ARCHIVE", "The archive contains no entries.");
  }
  if (declared.length > limits.maxEntries) {
    throw new ArchiveError(
      "ARCHIVE_LIMIT_EXCEEDED",
      `The archive contains ${declared.length} entries, more than the limit of ${limits.maxEntries}.`,
    );
  }

  let declaredTotal = 0;
  for (const entry of declared) {
    if (entry.originalSize > limits.maxEntryBytes) {
      throw new ArchiveError(
        "ARCHIVE_LIMIT_EXCEEDED",
        `Archive entry "${entry.name}" declares an uncompressed size above the per-entry limit.`,
      );
    }
    declaredTotal += entry.originalSize;
  }
  if (declaredTotal > limits.maxTotalBytes) {
    throw new ArchiveError(
      "ARCHIVE_LIMIT_EXCEEDED",
      "The archive declares more uncompressed data than the configured total limit.",
    );
  }
  if (bytes.length > 0 && declaredTotal / bytes.length > limits.maxCompressionRatio) {
    throw new ArchiveError(
      "ARCHIVE_LIMIT_EXCEEDED",
      "The archive expansion ratio is above the configured limit (possible decompression bomb).",
    );
  }

  const wanted = new Set<string>();
  for (const entry of declared) {
    if (isSafeEntryName(entry.name) && isUsefulPart(entry.name)) {
      wanted.add(normalizeName(entry.name).toLowerCase());
    }
  }

  let extracted: Record<string, Uint8Array>;
  try {
    extracted = unzipSync(bytes, {
      filter: (file: UnzipFileInfo) => wanted.has(normalizeName(file.name).toLowerCase()),
    });
  } catch (error) {
    throw new ArchiveError(
      "CORRUPT_ARCHIVE",
      "The archive contents could not be decompressed.",
      error instanceof Error ? error.message : String(error),
    );
  }

  const entries = new Map<string, ArchiveEntry>();
  let actualTotal = 0;
  const names: string[] = [];
  for (const entry of declared) {
    const normalized = normalizeName(entry.name);
    names.push(normalized);
    const key = normalized.toLowerCase();
    if (!wanted.has(key)) continue;
    const data = extracted[entry.name] ?? extracted[normalized];
    if (!data) continue;
    // The declared sizes are untrusted, so the real, inflated sizes are checked.
    if (data.byteLength > limits.maxEntryBytes) {
      throw new ArchiveError(
        "ARCHIVE_LIMIT_EXCEEDED",
        `Archive entry "${normalized}" expanded beyond the per-entry limit.`,
      );
    }
    actualTotal += data.byteLength;
    if (actualTotal > limits.maxTotalBytes) {
      throw new ArchiveError(
        "ARCHIVE_LIMIT_EXCEEDED",
        "Extracted archive content exceeded the configured total limit.",
      );
    }
    entries.set(key, { name: normalized, bytes: data });
  }

  return { entries, names, limits };
}

export function readBinary(archive: Archive, partName: string): Uint8Array | null {
  const key = normalizeName(partName).toLowerCase();
  return archive.entries.get(key)?.bytes ?? null;
}

const decoder = new TextDecoder("utf-8");

export function readText(archive: Archive, partName: string): string | null {
  const bytes = readBinary(archive, partName);
  if (!bytes) return null;
  if (bytes.byteLength > archive.limits.maxXmlBytes) {
    throw new XmlError(`XML part "${partName}" exceeds the configured size limit`);
  }
  let text = decoder.decode(bytes);
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1); // strip BOM
  return text;
}

export function hasPart(archive: Archive, partName: string): boolean {
  return archive.entries.has(normalizeName(partName).toLowerCase());
}

export function listParts(archive: Archive, prefix: string): string[] {
  const lower = prefix.toLowerCase();
  const out: string[] = [];
  for (const entry of archive.entries.values()) {
    if (entry.name.toLowerCase().startsWith(lower)) out.push(entry.name);
  }
  out.sort();
  return out;
}

/* ------------------------------------------------------- content types map */

export interface ContentTypes {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

export function parseContentTypes(root: import("./xml").XmlNode | null): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  if (!root) return { defaults, overrides };
  for (const node of root.children) {
    if (node.local === "Default") {
      const extension = node.attrs["Extension"];
      const type = node.attrs["ContentType"];
      if (extension && type) defaults.set(extension.toLowerCase(), type);
    } else if (node.local === "Override") {
      const partName = node.attrs["PartName"];
      const type = node.attrs["ContentType"];
      if (partName && type) overrides.set(normalizeName(partName).toLowerCase(), type);
    }
  }
  return { defaults, overrides };
}

export function contentTypeFor(types: ContentTypes, partName: string): string | null {
  const key = normalizeName(partName).toLowerCase();
  const override = types.overrides.get(key);
  if (override) return override;
  const dot = key.lastIndexOf(".");
  if (dot === -1) return null;
  return types.defaults.get(key.slice(dot + 1)) ?? null;
}

/* ----------------------------------------------------------- image sniffing */

export type ImageKind =
  | "png"
  | "jpeg"
  | "gif"
  | "bmp"
  | "webp"
  | "avif"
  | "svg"
  | "tiff"
  | "emf"
  | "wmf"
  | "unknown";

const BROWSER_RENDERABLE: ReadonlySet<ImageKind> = new Set<ImageKind>([
  "png",
  "jpeg",
  "gif",
  "bmp",
  "webp",
  "avif",
  "svg",
]);

export function isBrowserRenderable(kind: ImageKind): boolean {
  return BROWSER_RENDERABLE.has(kind);
}

const MIME_BY_KIND: Record<ImageKind, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
  avif: "image/avif",
  svg: "image/svg+xml",
  tiff: "image/tiff",
  emf: "image/emf",
  wmf: "image/wmf",
  unknown: "application/octet-stream",
};

export function mimeForKind(kind: ImageKind): string {
  return MIME_BY_KIND[kind];
}

function hasAsciiAt(bytes: Uint8Array, offset: number, text: string): boolean {
  if (bytes.length < offset + text.length) return false;
  for (let i = 0; i < text.length; i += 1) {
    if (bytes[offset + i] !== text.charCodeAt(i)) return false;
  }
  return true;
}

/**
 * Identify an image from its magic bytes.
 *
 * The declared content type is deliberately not trusted: a part can claim to be
 * a PNG while carrying something else, and the renderer needs to decide whether
 * the browser will be able to decode it at all.
 */
export function sniffImageKind(bytes: Uint8Array): ImageKind {
  if (bytes.length >= 8 && bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
    return "png";
  }
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "jpeg";
  if (hasAsciiAt(bytes, 0, "GIF8")) return "gif";
  if (bytes.length >= 2 && bytes[0] === 0x42 && bytes[1] === 0x4d) return "bmp";
  if (hasAsciiAt(bytes, 0, "RIFF") && hasAsciiAt(bytes, 8, "WEBP")) return "webp";
  // AVIF/AVIF-sequence: an ISO-BMFF file whose major or compatible brand is
  // `avif`/`avis`. The compatible brands follow the first one, so both positions
  // are checked.
  if (
    hasAsciiAt(bytes, 4, "ftyp") &&
    (hasAsciiAt(bytes, 8, "avif") || hasAsciiAt(bytes, 8, "avis") || hasAsciiAt(bytes, 16, "avif"))
  ) {
    return "avif";
  }
  if (
    bytes.length >= 4 &&
    ((bytes[0] === 0x49 && bytes[1] === 0x49 && bytes[2] === 0x2a && bytes[3] === 0x00) ||
      (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0x00 && bytes[3] === 0x2a))
  ) {
    return "tiff";
  }
  if (bytes.length >= 4 && bytes[0] === 0x01 && bytes[1] === 0x00 && bytes[2] === 0x00 && bytes[3] === 0x00) {
    if (hasAsciiAt(bytes, 40, " EMF")) return "emf";
  }
  if (bytes.length >= 4 && bytes[0] === 0xd7 && bytes[1] === 0xcd && bytes[2] === 0xc6 && bytes[3] === 0x9a) {
    return "wmf";
  }
  // SVG is text; look for an `<svg` element in the first kilobyte.
  const head = bytes.subarray(0, Math.min(bytes.length, 1024));
  let asText = "";
  for (const byte of head) asText += String.fromCharCode(byte);
  if (/<svg[\s>]/i.test(asText)) return "svg";
  return "unknown";
}
