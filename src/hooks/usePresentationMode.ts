import { useCallback, useEffect, useState } from "react";

export interface PresentationMode {
  /** true while the chrome-free slideshow is running */
  active: boolean;
  /** true when the control should be usable (a presentation is loaded) */
  available: boolean;
  enter: () => void;
  exit: () => void;
  toggle: () => void;
}

/**
 * Presentation (slideshow) mode.
 *
 * Works controlled (`presentation`) or uncontrolled (`defaultPresentation`), like
 * the slide index and zoom. The state lives outside the render tree so a host app
 * can drive it from its own toolbar, and so leaving presentation mode restores the
 * chrome exactly as it was (thumbnails, diagnostics and zoom are untouched).
 */
export function usePresentationMode(options: {
  controlled: boolean | undefined;
  initial: boolean;
  /** false while nothing is loaded; presenting is not possible then */
  available: boolean;
  onActiveChange?: (active: boolean) => void;
}): PresentationMode {
  const { controlled, initial, available, onActiveChange } = options;
  const [internal, setInternal] = useState(initial);

  const active = controlled ?? internal;

  const change = useCallback(
    (next: boolean) => {
      if (controlled === undefined) setInternal(next);
      if (next !== (controlled ?? internal)) onActiveChange?.(next);
    },
    [controlled, internal, onActiveChange],
  );

  const enter = useCallback(() => change(true), [change]);
  const exit = useCallback(() => change(false), [change]);
  const toggle = useCallback(() => change(!active), [change, active]);

  // A presentation that fails to load, or a new source, must not stay in a
  // slideshow that has nothing to show.
  useEffect(() => {
    if (!available && active) change(false);
  }, [available, active, change]);

  return { active, available, enter, exit, toggle };
}
