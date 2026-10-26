/**
 * Raster image fixtures.
 *
 * A PNG fixture already lives with the presentation fixtures because decks embed
 * one. The demo's browser test also needs the other bitmap types a file picker
 * advertises, so they are generated here (BMP) or embedded as a tiny real file
 * (JPEG) rather than being checked in as binaries.
 */

/**
 * Build an uncompressed 24-bit BMP.
 *
 * Rows are padded to four bytes, as the format requires, so the fixture also
 * exercises the header reader's bottom-up row handling rather than only the
 * magic bytes.
 */
export function buildBmp(width = 4, height = 3, rgb: [number, number, number] = [30, 120, 200]): Uint8Array {
  const rowBytes = Math.ceil((width * 3) / 4) * 4;
  const pixelBytes = rowBytes * height;
  const offset = 54;
  const bytes = new Uint8Array(offset + pixelBytes);
  const view = new DataView(bytes.buffer);

  // BITMAPFILEHEADER
  bytes[0] = 0x42; // B
  bytes[1] = 0x4d; // M
  view.setUint32(2, bytes.length, true);
  view.setUint32(10, offset, true);

  // BITMAPINFOHEADER
  view.setUint32(14, 40, true);
  view.setInt32(18, width, true);
  view.setInt32(22, height, true);
  view.setUint16(26, 1, true);
  view.setUint16(28, 24, true);
  view.setUint32(34, pixelBytes, true);
  view.setInt32(38, 2835, true);
  view.setInt32(42, 2835, true);

  const [r, g, b] = rgb;
  for (let row = 0; row < height; row += 1) {
    // BMP rows are stored bottom-up.
    const at = offset + (height - 1 - row) * rowBytes;
    for (let column = 0; column < width; column += 1) {
      const pixel = at + column * 3;
      bytes[pixel] = b;
      bytes[pixel + 1] = g;
      bytes[pixel + 2] = r;
    }
  }
  return bytes;
}

/**
 * A real 8×8 baseline JPEG, produced once from a generated PNG. Embedding the
 * bytes keeps the fixture deterministic; it is the smallest honest JPEG the
 * browser can decode, not a hand-shaped header.
 */
export const TINY_JPEG = Uint8Array.from(
  atob(
    "/9j/4AAQSkZJRgABAQAASABIAAD/4QBMRXhpZgAATU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAAACKADAAQAAAABAAAACAAAAAD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZjwCyBOmACZjs+EJ+/8AAEQgACAAIAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/bAEMAAgICAgICAwICAwUDAwMFBgUFBQUGCAYGBgYGCAoICAgICAgKCgoKCgoKCgwMDAwMDA4ODg4ODw8PDw8PDw8PD//bAEMBAgICBAQEBwQEBxALCQsQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEP/dAAQAAf/aAAwDAQACEQMRAD8A+Q6KKK/Cz/VQ/9k=",
  ),
  (character) => character.charCodeAt(0),
);
