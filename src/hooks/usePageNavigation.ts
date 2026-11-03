import { useCallback, useEffect, useRef, useState, type MutableRefObject } from "react";

function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value;
}

export interface UsePageNavigationOptions {
  /** pages in the loaded document; 0 while nothing is loaded */
  count: number;
  /** the `slide` prop, when the host controls the index */
  controlled: number | undefined;
  /** the `defaultSlide` prop, also restored on every new load */
  defaultIndex: number;
  /** changes whenever a new document has loaded (the viewer's `loadId`) */
  resetKey: number;
  onChange?: (index: number) => void;
}

export interface PageNavigation {
  /** the page on screen, always inside the document's bounds */
  index: number;
  /** current values for callbacks that must not be recreated per render */
  indexRef: MutableRefObject<number>;
  countRef: MutableRefObject<number>;
  goTo: (index: number) => void;
  next: () => void;
  previous: () => void;
}

/**
 * The current page, controlled or not.
 *
 * Every navigation callback is stable and reads the live index and count from
 * refs, so the toolbar, keyboard handlers and the imperative handle never see a
 * stale page or need to be rebuilt when the page changes.
 */
export function usePageNavigation({
  count,
  controlled,
  defaultIndex,
  resetKey,
  onChange,
}: UsePageNavigationOptions): PageNavigation {
  const [internalIndex, setInternalIndex] = useState(defaultIndex);
  const requestedIndex = controlled ?? internalIndex;
  const index = count === 0 ? 0 : clamp(Math.round(requestedIndex), 0, count - 1);

  const indexRef = useRef(index);
  indexRef.current = index;
  const countRef = useRef(count);
  countRef.current = count;
  const controlledRef = useRef(controlled);
  controlledRef.current = controlled;
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;

  const goTo = useCallback((target: number) => {
    const currentCount = countRef.current;
    if (currentCount === 0) return;
    const next = clamp(Math.round(target), 0, currentCount - 1);
    if (controlledRef.current === undefined) setInternalIndex(next);
    if (next !== indexRef.current) onChangeRef.current?.(next);
  }, []);

  const next = useCallback(() => goTo(indexRef.current + 1), [goTo]);
  const previous = useCallback(() => goTo(indexRef.current - 1), [goTo]);

  // One effect, because the two corrections must not race: a newly loaded
  // document starts again from the default page, and otherwise an out-of-range
  // (or controlled) index is pulled back inside the document's bounds. Either way
  // the host hears the page that is actually shown — a reset must not leave it
  // with the clamped index of the previous document.
  const lastReset = useRef({ key: resetKey, defaultIndex });
  useEffect(() => {
    const reset = lastReset.current.key !== resetKey || lastReset.current.defaultIndex !== defaultIndex;
    lastReset.current = { key: resetKey, defaultIndex };
    if (reset && resetKey !== 0) {
      setInternalIndex(defaultIndex);
      if (controlled === undefined) {
        const target = count === 0 ? 0 : clamp(Math.round(defaultIndex), 0, count - 1);
        if (target !== internalIndex) onChangeRef.current?.(target);
        return;
      }
    }
    if (count === 0) return;
    if (requestedIndex < 0 || requestedIndex > count - 1 || !Number.isFinite(requestedIndex)) {
      setInternalIndex(index);
      onChangeRef.current?.(index);
    }
  }, [controlled, count, defaultIndex, index, internalIndex, requestedIndex, resetKey]);

  return { index, indexRef, countRef, goTo, next, previous };
}
