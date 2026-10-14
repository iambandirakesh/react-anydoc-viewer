import { useCallback, useEffect, useState } from "react";

export interface ElementSize {
  width: number;
  height: number;
  /** false until a non-zero measurement has been observed */
  measured: boolean;
}

export type MeasuredRef<T extends HTMLElement> = (node: T | null) => void;

/**
 * Track an element's size.
 *
 * A callback ref is used rather than a `RefObject` so that elements which mount
 * later (the slide canvas appears only once a presentation is parsed) are still
 * measured. A zero-sized container is a common cause of an apparently blank
 * viewer, so `measured` is exposed for the viewer to explain that state instead
 * of rendering nothing.
 */
export function useElementSize<T extends HTMLElement>(): [MeasuredRef<T>, ElementSize] {
  const [node, setNode] = useState<T | null>(null);
  const [size, setSize] = useState<ElementSize>({ width: 0, height: 0, measured: false });

  const attach = useCallback((next: T | null) => {
    setNode(next);
  }, []);

  useEffect(() => {
    if (!node) return undefined;

    const update = (width: number, height: number) => {
      setSize((previous) => {
        const roundedWidth = Math.round(width);
        const roundedHeight = Math.round(height);
        if (previous.width === roundedWidth && previous.height === roundedHeight && previous.measured) {
          return previous;
        }
        return { width: roundedWidth, height: roundedHeight, measured: roundedWidth > 0 && roundedHeight > 0 };
      });
    };

    const rect = node.getBoundingClientRect();
    update(rect.width, rect.height);

    if (typeof ResizeObserver !== "undefined") {
      const observer = new ResizeObserver((entries) => {
        for (const entry of entries) {
          const box = entry.contentRect;
          update(box.width, box.height);
        }
      });
      observer.observe(node);
      return () => observer.disconnect();
    }

    const onWindowResize = () => {
      const next = node.getBoundingClientRect();
      update(next.width, next.height);
    };
    window.addEventListener("resize", onWindowResize);
    return () => window.removeEventListener("resize", onWindowResize);
  }, [node]);

  return [attach, size];
}
