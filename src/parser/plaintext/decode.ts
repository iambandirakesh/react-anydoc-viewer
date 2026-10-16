/**
 * Text decoding for the plain-text formats.
 *
 * A byte-order mark is the only reliable signal a text file carries, so it is
 * checked first and then the three realistic cases are tried in order: strict
 * UTF-8, UTF-16 without a mark (recognised from its NUL pattern), and finally
 * Windows-1252 (which decodes every byte and so can never fail — the honest
 * choice for "some legacy encoding").
 *
 * The chosen encoding is always reported, because silently replacing characters
 * is how mojibake reaches a user without anyone noticing.
 */

export type TextEncoding = "utf-8" | "utf-8-bom" | "utf-16le" | "utf-16be" | "windows-1252";

export interface DecodedText {
  text: string;
  encoding: TextEncoding;
  /** true when the fallback was used because strict UTF-8 rejected the bytes */
  lossy: boolean;
}

/** BOMs, longest first so a UTF-8 mark is not mistaken for anything else. */
const BOMS: ReadonlyArray<{ bytes: readonly number[]; encoding: TextEncoding }> = [
  { bytes: [0xef, 0xbb, 0xbf], encoding: "utf-8-bom" },
  { bytes: [0xff, 0xfe], encoding: "utf-16le" },
  { bytes: [0xfe, 0xff], encoding: "utf-16be" },
];

function startsWith(bytes: Uint8Array, mark: readonly number[]): boolean {
  if (bytes.length < mark.length) return false;
  for (let i = 0; i < mark.length; i += 1) {
    if (bytes[i] !== mark[i]) return false;
  }
  return true;
}

function decodeWith(bytes: Uint8Array, label: string, fatal: boolean): string | null {
  try {
    return new TextDecoder(label, { fatal }).decode(bytes);
  } catch {
    return null;
  }
}

/**
 * Heuristic for UTF-16 without a BOM: ASCII text stored as UTF-16LE is half zero
 * bytes, and those zeros sit in the second byte of every pair.
 */
function looksLikeUtf16(bytes: Uint8Array): "utf-16le" | "utf-16be" | null {
  const sample = Math.min(bytes.length, 4096);
  // Two code units are enough to see the pattern; a one-byte file cannot be
  // UTF-16 at all.
  if (sample < 4) return null;
  let oddZeros = 0;
  let evenZeros = 0;
  for (let i = 0; i < sample; i += 2) {
    if (bytes[i] === 0) evenZeros += 1;
    if (bytes[i + 1] === 0) oddZeros += 1;
  }
  const pairs = Math.floor(sample / 2);
  if (oddZeros / pairs > 0.3 && evenZeros / pairs < 0.05) return "utf-16le";
  if (evenZeros / pairs > 0.3 && oddZeros / pairs < 0.05) return "utf-16be";
  return null;
}

export function decodeTextFile(bytes: Uint8Array): DecodedText {
  for (const bom of BOMS) {
    if (!startsWith(bytes, bom.bytes)) continue;
    const body = bytes.subarray(bom.bytes.length);
    // A UTF-16 file with a mark still carries the mark's endianness.
    const label = bom.encoding === "utf-16le" ? "utf-16le" : bom.encoding === "utf-16be" ? "utf-16be" : "utf-8";
    const text = decodeWith(body, label, false) ?? "";
    return { text, encoding: bom.encoding, lossy: false };
  }

  // A NUL is not text: it is what UTF-16 without a mark looks like when read as
  // single-byte text, so its presence makes the UTF-16 shape worth testing even
  // though strict UTF-8 accepted the bytes.
  const strictUtf8 = decodeWith(bytes, "utf-8", true);
  if (strictUtf8 !== null && !strictUtf8.includes("\u0000")) {
    return { text: strictUtf8, encoding: "utf-8", lossy: false };
  }

  const wide = looksLikeUtf16(bytes);
  if (wide) {
    const text = decodeWith(bytes, wide, false);
    if (text !== null) return { text, encoding: wide, lossy: false };
  }
  if (strictUtf8 !== null) return { text: strictUtf8, encoding: "utf-8", lossy: false };

  // Windows-1252 decodes every byte sequence, so this always succeeds.
  const fallback = decodeWith(bytes, "windows-1252", false) ?? "";
  return { text: fallback, encoding: "windows-1252", lossy: true };
}

/**
 * How much of a text file looks like text at all.
 *
 * Used by detection, which must not claim a binary blob is a text file. NUL
 * bytes and a high proportion of control characters are the signals.
 */
export function looksLikeText(bytes: Uint8Array, sampleSize = 4096): boolean {
  const sample = bytes.subarray(0, Math.min(bytes.length, sampleSize));
  if (sample.length === 0) return false;
  const decoded = decodeTextFile(sample);
  let controls = 0;
  let printable = 0;
  for (let i = 0; i < decoded.text.length; i += 1) {
    const code = decoded.text.charCodeAt(i);
    // Tab and newline are ordinary text; everything else below 0x20 is not.
    if (code === 0x09 || code === 0x0a || code === 0x0d) continue;
    if (code < 0x20 || code === 0x7f) controls += 1;
    else if (code !== 0x20) printable += 1;
  }
  // Whitespace alone is not a document: a file of nothing but tabs or spaces has
  // no text to show, and calling it a text file would turn binary padding into a
  // blank page.
  return controls === 0 && printable > 0;
}
