/**
 * Minimal declarations for `utif` (MIT, no bundled types).
 *
 * Only the three calls the TIFF page renderer uses are declared. The real module
 * also exposes EXIF/RAW helpers that the viewer does not need, so declaring the
 * whole surface would be a promise the package cannot keep.
 */
declare module "utif" {
  export function decode(buffer: ArrayBuffer): unknown[];
  export function decodeImage(buffer: ArrayBuffer, ifd: unknown): void;
  export function toRGBA8(ifd: unknown): Uint8Array;
  const UTIF: {
    decode: typeof decode;
    decodeImage: typeof decodeImage;
    toRGBA8: typeof toRGBA8;
  };
  export default UTIF;
}
