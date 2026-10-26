/**
 * TIFF fixture builder.
 *
 * The images the test suite uses are generated rather than checked in, matching
 * the `.pptx`, `.docx` and `.pdf` builders. Each fixture is a classic little-
 * endian TIFF with uncompressed RGB strips and a real IFD chain, so it exercises
 * the same directory walk and decoder path as a file from a camera or scanner.
 */

export interface TiffPageSpec {
  width: number;
  height: number;
  /** solid RGB colour; defaults to red */
  rgb?: [number, number, number];
}

export interface TiffFixtureOptions {
  pages?: TiffPageSpec[];
}

const IFD_ENTRIES = 10;
const IFD_SIZE = 2 + IFD_ENTRIES * 12 + 4 + 6; // entries + next offset + 3 shorts

function writeUint16(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >> 8) & 0xff;
}

function writeUint32(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >> 8) & 0xff;
  target[offset + 2] = (value >> 16) & 0xff;
  target[offset + 3] = (value >> 24) & 0xff;
}

function writeEntry(
  target: Uint8Array,
  ifdOffset: number,
  index: number,
  tag: number,
  type: number,
  count: number,
  value: number,
): void {
  const entry = ifdOffset + 2 + index * 12;
  writeUint16(target, entry, tag);
  writeUint16(target, entry + 2, type);
  writeUint32(target, entry + 4, count);
  if (type === 3 && count === 1) writeUint16(target, entry + 8, value);
  else writeUint32(target, entry + 8, value);
}

/**
 * Build a classic uncompressed RGB TIFF.
 *
 * Layout: header, one 132-byte IFD per page (including the three BitsPerSample
 * shorts), then each page's pixel strip. Every offset is derived from that layout,
 * so adding a page only moves later offsets.
 */
export function buildTiff(options: TiffFixtureOptions = {}): Uint8Array {
  const pages = options.pages ?? [{ width: 2, height: 2 }];
  const ifdBase = 8;
  const pixelBase = ifdBase + pages.length * IFD_SIZE;
  const totalPixels = pages.reduce((sum, page) => sum + page.width * page.height * 3, 0);
  const bytes = new Uint8Array(pixelBase + totalPixels);

  // Classic little-endian TIFF header.
  bytes[0] = 0x49; // I
  bytes[1] = 0x49; // I
  writeUint16(bytes, 2, 42);
  writeUint32(bytes, 4, ifdBase);

  let pixelOffset = pixelBase;
  pages.forEach((page, pageIndex) => {
    const ifdOffset = ifdBase + pageIndex * IFD_SIZE;
    const bitsOffset = ifdOffset + 2 + IFD_ENTRIES * 12 + 4;
    const stripBytes = page.width * page.height * 3;

    writeUint16(bytes, ifdOffset, IFD_ENTRIES);
    writeEntry(bytes, ifdOffset, 0, 256, 4, 1, page.width);
    writeEntry(bytes, ifdOffset, 1, 257, 4, 1, page.height);
    writeEntry(bytes, ifdOffset, 2, 258, 3, 3, bitsOffset);
    writeEntry(bytes, ifdOffset, 3, 259, 3, 1, 1); // no compression
    writeEntry(bytes, ifdOffset, 4, 262, 3, 1, 2); // RGB
    writeEntry(bytes, ifdOffset, 5, 273, 4, 1, pixelOffset);
    writeEntry(bytes, ifdOffset, 6, 277, 3, 1, 3); // three samples per pixel
    writeEntry(bytes, ifdOffset, 7, 278, 4, 1, page.height);
    writeEntry(bytes, ifdOffset, 8, 279, 4, 1, stripBytes);
    writeEntry(bytes, ifdOffset, 9, 284, 3, 1, 1); // chunky
    writeUint32(bytes, ifdOffset + 2 + IFD_ENTRIES * 12, pageIndex + 1 < pages.length ? ifdOffset + IFD_SIZE : 0);

    // BitsPerSample: 8, 8, 8.
    writeUint16(bytes, bitsOffset, 8);
    writeUint16(bytes, bitsOffset + 2, 8);
    writeUint16(bytes, bitsOffset + 4, 8);

    const [r, g, b] = page.rgb ?? [255, 0, 0];
    for (let pixel = 0; pixel < page.width * page.height; pixel += 1) {
      const at = pixelOffset + pixel * 3;
      bytes[at] = r;
      bytes[at + 1] = g;
      bytes[at + 2] = b;
    }
    pixelOffset += stripBytes;
  });

  return bytes;
}

/** A header-only TIFF, for exercising the malformed-directory fallback. */
export function brokenTiff(): Uint8Array {
  const bytes = buildTiff({ pages: [{ width: 2, height: 2 }] }).slice(0, 8);
  return bytes;
}
