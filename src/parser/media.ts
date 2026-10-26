/**
 * Media store.
 *
 * Image bytes stay in the archive until a slide actually needs them, and the
 * resulting object URLs are tracked so they can be revoked on unmount. When
 * `URL.createObjectURL` is unavailable (server rendering, jsdom) we fall back to
 * `data:` URLs, which are also what keeps the component usable without a DOM.
 *
 * SVG parts are always delivered through an `<img>`-compatible URL: browsers do
 * not execute scripts in images, so untrusted SVG cannot reach the host page.
 */

import {
  DEFAULT_ARCHIVE_LIMITS,
  mimeForKind,
  readBinary,
  sniffImageKind,
  type Archive,
  type ImageKind,
} from "./archive";

export interface MediaStore {
  /** image format detected from the bytes, or null when the part is absent */
  kind(partName: string): ImageKind | null;
  /** URL usable as an `<img src>`, or null when the part is missing/undecodable */
  url(partName: string): string | null;
  mime(partName: string): string;
  byteLength(partName: string): number;
  /** revoke every object URL created so far */
  dispose(): void;
  /** number of object URLs currently held */
  activeUrls: number;
}

/**
 * A URL that hands `bytes` to the browser as an image of `mime`.
 *
 * Object URLs are used when they exist (they keep the bytes out of the DOM and
 * out of the JS heap), and a `data:` URL otherwise, which is what keeps the
 * component usable without a DOM. `null` means neither was available.
 */
export function createByteUrl(bytes: Uint8Array, mime: string): { url: string; revoke(): void } | null {
  if (typeof URL !== "undefined" && typeof URL.createObjectURL === "function" && typeof Blob !== "undefined") {
    const url = URL.createObjectURL(new Blob([bytes as unknown as BlobPart], { type: mime }));
    return {
      url,
      revoke: () => {
        try {
          URL.revokeObjectURL(url);
        } catch {
          /* already revoked */
        }
      },
    };
  }
  const base64 = toBase64(bytes);
  if (base64.length === 0) return null;
  return { url: `data:${mime};base64,${base64}`, revoke: () => undefined };
}

function toBase64(bytes: Uint8Array): string {
  const chunk = 0x8000;
  let binary = "";
  for (let i = 0; i < bytes.length; i += chunk) {
    const slice = bytes.subarray(i, Math.min(i + chunk, bytes.length));
    let part = "";
    for (let j = 0; j < slice.length; j += 1) part += String.fromCharCode(slice[j] as number);
    binary += part;
  }
  if (typeof btoa === "function") return btoa(binary);
  // Node (tests, SSR): Buffer is available without a DOM.
  const buffer = (globalThis as { Buffer?: { from(input: string, encoding: string): { toString(enc: string): string } } })
    .Buffer;
  if (buffer) return buffer.from(binary, "binary").toString("base64");
  return "";
}

/**
 * A media store with no parts at all.
 *
 * Formats whose pages are not drawn from a shape list (PDF) still render through
 * the same components, which expect a store; this is that empty one.
 */
export function createEmptyMediaStore(): MediaStore {
  return createMediaStore({ entries: new Map(), names: [], limits: DEFAULT_ARCHIVE_LIMITS });
}

export function createMediaStore(archive: Archive): MediaStore {
  const kinds = new Map<string, ImageKind | null>();
  const urls = new Map<string, string>();
  const sizes = new Map<string, number>();
  let created: string[] = [];

  const canUseObjectUrl =
    typeof URL !== "undefined" &&
    typeof URL.createObjectURL === "function" &&
    typeof Blob !== "undefined";


  function kindOf(partName: string): ImageKind | null {
    const key = partName.toLowerCase();
    if (kinds.has(key)) return kinds.get(key) ?? null;
    const bytes = readBinary(archive, partName);
    const kind = bytes ? sniffImageKind(bytes) : null;
    kinds.set(key, kind);
    if (bytes) sizes.set(key, bytes.byteLength);
    return kind;
  }

  return {
    kind: kindOf,
    mime(partName) {
      const kind = kindOf(partName);
      return kind ? mimeForKind(kind) : "application/octet-stream";
    },
    byteLength(partName) {
      const key = partName.toLowerCase();
      if (!sizes.has(key)) kindOf(partName);
      return sizes.get(key) ?? 0;
    },
    url(partName) {
      const key = partName.toLowerCase();
      const existing = urls.get(key);
      if (existing) return existing;
      const bytes = readBinary(archive, partName);
      if (!bytes) return null;
      const kind = kindOf(partName);
      if (!kind) return null;
      const mime = mimeForKind(kind);
      const stored = createByteUrl(bytes, mime);
      if (!stored) return null;
      if (canUseObjectUrl) created.push(stored.url);
      urls.set(key, stored.url);
      return stored.url;
    },
    dispose() {
      if (canUseObjectUrl) {
        for (const url of created) {
          try {
            URL.revokeObjectURL(url);
          } catch {
            /* already revoked */
          }
        }
      }
      created = [];
      urls.clear();
    },
    get activeUrls() {
      return canUseObjectUrl ? created.length : 0;
    },
  };
}
