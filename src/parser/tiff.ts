/**
 * TIFF directory inspection.
 *
 * A standalone TIFF cannot be handed to an `<img>` element: Chrome and Firefox
 * do not decode it, so the viewer has to decode the pixels itself. That decode
 * lives behind an `import()` in `TiffPage`, because it is only needed when a TIFF
 * is actually opened; this module only walks the IFD chain to learn how many
 * pages exist and what their natural sizes are. It never reads pixel data.
 *
 * The parser is deliberately bounded. A TIFF header can point an IFD anywhere,
 * including back at itself, so offsets are validated against the input and the
 * page count is capped. Anything suspicious is treated as "no usable directory"
 * and the caller falls back to one default page instead of looping forever.
 */

export interface TiffPageInfo {
  /** natural width in CSS pixels */
  width: number;
  /** natural height in CSS pixels */
  height: number;
}

/** TIFF tags this module needs. */
const TAG_IMAGE_WIDTH = 256;
const TAG_IMAGE_LENGTH = 257;

const TYPE_SHORT = 3;
const TYPE_LONG = 4;

/** Refuse absurd dimensions rather than allocating a canvas around them. */
const MAX_DIMENSION = 20_000;
/** A page count beyond this is not a document a person is reading. */
const MAX_PAGES = 100;

function readUint16(bytes: Uint8Array, offset: number, littleEndian: boolean): number {
  if (offset + 2 > bytes.length) return NaN;
  return littleEndian
    ? (bytes[offset] as number) | ((bytes[offset + 1] as number) << 8)
    : ((bytes[offset] as number) << 8) | (bytes[offset + 1] as number);
}

function readUint32(bytes: Uint8Array, offset: number, littleEndian: boolean): number {
  if (offset + 4 > bytes.length) return NaN;
  return littleEndian
    ? (bytes[offset] as number) |
        ((bytes[offset + 1] as number) << 8) |
        ((bytes[offset + 2] as number) << 16) |
        ((bytes[offset + 3] as number) << 24)
    : ((bytes[offset] as number) << 24) |
        ((bytes[offset + 1] as number) << 16) |
        ((bytes[offset + 2] as number) << 8) |
        (bytes[offset + 3] as number);
}

function readValue(
  bytes: Uint8Array,
  entryOffset: number,
  type: number,
  count: number,
  littleEndian: boolean,
): number | null {
  if (count < 1) return null;
  if (type === TYPE_SHORT) {
    return readUint16(bytes, entryOffset + 8, littleEndian);
  }
  if (type === TYPE_LONG) {
    return readUint32(bytes, entryOffset + 8, littleEndian);
  }
  // Width and height are allowed to be SHORT or LONG; anything else is not a
  // geometry tag we can trust.
  return null;
}

/**
 * Walk the IFD (image file directory) chain and report every page's size.
 *
 * Returns an empty array when the header is not a classic TIFF, when the chain
 * is malformed, or when no directory carries a plausible size. The caller owns
 * the fallback because the fallback is a presentation decision, not a parsing one.
 */
export function readTiffPages(bytes: Uint8Array): TiffPageInfo[] {
  if (bytes.length < 8) return [];
  const order = String.fromCharCode(bytes[0] as number, bytes[1] as number);
  if (order !== "II" && order !== "MM") return [];
  const littleEndian = order === "II";
  const magic = readUint16(bytes, 2, littleEndian);
  if (magic !== 42) return [];

  const pages: TiffPageInfo[] = [];
  const seen = new Set<number>();
  let offset = readUint32(bytes, 4, littleEndian);
  if (!Number.isFinite(offset) || offset < 8 || offset >= bytes.length) return [];

  while (pages.length < MAX_PAGES) {
    if (seen.has(offset) || offset + 2 > bytes.length) break;
    seen.add(offset);

    const count = readUint16(bytes, offset, littleEndian);
    if (!Number.isFinite(count) || count === 0) break;
    const entriesStart = offset + 2;
    const entriesEnd = entriesStart + count * 12;
    const nextOffsetAt = entriesEnd;
    if (entriesEnd + 4 > bytes.length) break;

    let width: number | null = null;
    let height: number | null = null;
    for (let index = 0; index < count; index += 1) {
      const entry = entriesStart + index * 12;
      if (entry + 12 > bytes.length) break;
      const tag = readUint16(bytes, entry, littleEndian);
      const type = readUint16(bytes, entry + 2, littleEndian);
      const valueCount = readUint32(bytes, entry + 4, littleEndian);
      if (tag === TAG_IMAGE_WIDTH) width = readValue(bytes, entry, type, valueCount, littleEndian);
      else if (tag === TAG_IMAGE_LENGTH) height = readValue(bytes, entry, type, valueCount, littleEndian);
    }

    if (
      width !== null &&
      height !== null &&
      Number.isFinite(width) &&
      Number.isFinite(height) &&
      width > 0 &&
      height > 0 &&
      width <= MAX_DIMENSION &&
      height <= MAX_DIMENSION
    ) {
      pages.push({ width: Math.round(width), height: Math.round(height) });
    } else {
      // A directory without usable geometry still means one page, but the
      // renderer will letterbox it into a default box. Do not invent a size here.
      pages.push({ width: 0, height: 0 });
    }

    const next = readUint32(bytes, nextOffsetAt, littleEndian);
    if (!Number.isFinite(next) || next === 0) break;
    if (next < 8 || next >= bytes.length || seen.has(next)) break;
    offset = next;
  }

  return pages;
}
