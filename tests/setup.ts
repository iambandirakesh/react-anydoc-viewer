/**
 * Vitest setup.
 *
 * The viewer measures its container with `getBoundingClientRect` plus a
 * `ResizeObserver`. happy-dom returns a zero rect and never fires the observer,
 * which would make every component test see an unmeasured container, so the
 * environment is made deterministic here:
 *
 *  - `getBoundingClientRect` returns a fixed 900x600 box (or zero when
 *    {@link setZeroSize} is used, for the "no usable size" path);
 *  - `ResizeObserver` is removed so the hook uses its documented fallback;
 *  - object URLs are available because the media store prefers them.
 */

import "@testing-library/jest-dom/vitest";
import { afterEach, beforeEach, vi } from "vitest";

const DEFAULT_SIZE = { width: 900, height: 600 };

let zeroSize = false;

export function setZeroSize(value: boolean): void {
  zeroSize = value;
}

const created = new Set<string>();

beforeEach(() => {
  zeroSize = false;
  created.clear();

  const rect = () =>
    zeroSize
      ? {
          x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0,
          width: 0, height: 0, toJSON: () => ({}),
        }
      : {
          x: 0, y: 0, top: 0, left: 0, right: DEFAULT_SIZE.width, bottom: DEFAULT_SIZE.height,
          width: DEFAULT_SIZE.width, height: DEFAULT_SIZE.height, toJSON: () => ({}),
        };

  // A few suites run in the `node` environment (the PDF adapter needs no DOM),
  // where there is no element to patch.
  if (typeof HTMLElement !== "undefined") {
    Object.defineProperty(HTMLElement.prototype, "getBoundingClientRect", {
      configurable: true,
      writable: true,
      value: rect,
    });
  }

  if (typeof globalThis !== "undefined") {
    Object.defineProperty(globalThis, "ResizeObserver", {
      configurable: true,
      writable: true,
      value: undefined,
    });
  }

  if (typeof URL !== "undefined" && typeof URL.createObjectURL !== "function") {
    Object.defineProperty(URL, "createObjectURL", {
      configurable: true,
      writable: true,
      value: (blob: Blob) => {
        const url = `blob:mock/${created.size}-${blob.size}`;
        created.add(url);
        return url;
      },
    });
  }
  if (typeof URL !== "undefined" && typeof URL.revokeObjectURL !== "function") {
    Object.defineProperty(URL, "revokeObjectURL", {
      configurable: true,
      writable: true,
      value: () => undefined,
    });
  }
});

afterEach(() => {
  vi.restoreAllMocks();
});

export function createdObjectUrls(): string[] {
  return [...created];
}
