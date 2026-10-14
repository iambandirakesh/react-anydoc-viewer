import { useCallback, useEffect, useRef, useState, type RefObject } from "react";

interface FullscreenDocument extends Document {
  webkitFullscreenElement?: Element | null;
  webkitExitFullscreen?: () => Promise<void> | void;
}

interface FullscreenElement extends HTMLElement {
  webkitRequestFullscreen?: () => Promise<void> | void;
}

export interface FullscreenControls {
  /** true when either the native API or the overlay fallback is active */
  isFullscreen: boolean;
  /** true when the viewer is expanded with its own overlay instead of the API */
  usesOverlay: boolean;
  /** true when the native Fullscreen API exists and reports itself enabled */
  nativeSupported: boolean;
  enter: () => void;
  exit: () => void;
  toggle: () => void;
}

/**
 * How long to wait for `requestFullscreen()` to actually take effect.
 *
 * Some embedded environments (Electron-style webviews, iframes without an
 * `allowfullscreen` policy, kiosk browsers) advertise `document.fullscreenEnabled`
 * but never resolve the request nor fire `fullscreenchange`. Waiting briefly lets
 * the viewer fall back to its own overlay instead of leaving the button dead.
 */
const NATIVE_FULLSCREEN_TIMEOUT_MS = 450;

function nativeElement(doc: FullscreenDocument): Element | null {
  return doc.fullscreenElement ?? doc.webkitFullscreenElement ?? null;
}

/**
 * Fullscreen for a container element.
 *
 * Tries the real Fullscreen API first (with Safari's prefixed variant), and falls
 * back to a fixed-position overlay when the API is missing, blocked or silently
 * ignored. The fallback keeps `Esc`, the toolbar toggle and the `f` shortcut
 * working everywhere, which is the difference between a button that does nothing
 * and one that always fills the viewport.
 */
export function useFullscreen(target: RefObject<HTMLElement | null>): FullscreenControls {
  const [nativeActive, setNativeActive] = useState(false);
  const [overlayActive, setOverlayActive] = useState(false);
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const nativeSupported =
    typeof document !== "undefined" &&
    (document.fullscreenEnabled === true ||
      typeof (document as FullscreenDocument).webkitFullscreenElement !== "undefined");

  const clearTimer = useCallback(() => {
    if (timer.current !== null) {
      clearTimeout(timer.current);
      timer.current = null;
    }
  }, []);

  useEffect(() => {
    if (typeof document === "undefined") return undefined;
    const handler = () => {
      const element = nativeElement(document as FullscreenDocument);
      const entered = element !== null && element === target.current;
      setNativeActive(entered);
      if (entered) {
        clearTimer();
        setOverlayActive(false);
      }
    };
    document.addEventListener("fullscreenchange", handler);
    document.addEventListener("webkitfullscreenchange", handler);
    return () => {
      document.removeEventListener("fullscreenchange", handler);
      document.removeEventListener("webkitfullscreenchange", handler);
      clearTimer();
    };
  }, [target, clearTimer]);

  const enter = useCallback(() => {
    const element = target.current as FullscreenElement | null;
    if (!element) return;
    const request = element.requestFullscreen ?? element.webkitRequestFullscreen;
    if (!request) {
      setOverlayActive(true);
      return;
    }

    // Watchdog: if the browser neither enters fullscreen nor rejects, use the overlay.
    clearTimer();
    timer.current = setTimeout(() => {
      timer.current = null;
      if (nativeElement(document as FullscreenDocument) === null) setOverlayActive(true);
    }, NATIVE_FULLSCREEN_TIMEOUT_MS);

    try {
      const result = request.call(element);
      if (result && typeof (result as Promise<void>).then === "function") {
        (result as Promise<void>)
          .then(() => {
            // Chrome resolves before `fullscreenchange`; the watchdog decides.
            if (nativeElement(document as FullscreenDocument) === null) {
              clearTimer();
              setOverlayActive(true);
            }
          })
          .catch((error: unknown) => {
            clearTimer();
            setOverlayActive(true);
            if (typeof console !== "undefined") {
              console.warn(
                "[react-anydoc-viewer] requestFullscreen() was rejected; using the in-page fullscreen overlay instead.",
                error,
              );
            }
          });
      }
    } catch (error) {
      clearTimer();
      setOverlayActive(true);
      if (typeof console !== "undefined") {
        console.warn(
          "[react-anydoc-viewer] requestFullscreen() threw; using the in-page fullscreen overlay instead.",
          error,
        );
      }
    }
  }, [target, clearTimer]);

  const exit = useCallback(() => {
    setOverlayActive(false);
    if (typeof document === "undefined") return;
    const doc = document as FullscreenDocument;
    if (nativeElement(doc) === null) return;
    const exitFn = doc.exitFullscreen ?? doc.webkitExitFullscreen;
    if (!exitFn) return;
    try {
      const result = exitFn.call(doc);
      if (result && typeof (result as Promise<void>).catch === "function") {
        (result as Promise<void>).catch(() => undefined);
      }
    } catch {
      /* nothing else to do: the viewer is already out of native fullscreen */
    }
  }, []);

  const isFullscreen = nativeActive || overlayActive;

  const toggle = useCallback(() => {
    if (isFullscreen) exit();
    else enter();
  }, [enter, exit, isFullscreen]);

  // Esc leaves the overlay. Native fullscreen is handled by the browser itself.
  useEffect(() => {
    if (!overlayActive || typeof document === "undefined") return undefined;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.stopPropagation();
        setOverlayActive(false);
      }
    };
    document.addEventListener("keydown", onKeyDown, true);
    return () => document.removeEventListener("keydown", onKeyDown, true);
  }, [overlayActive]);

  return { isFullscreen, usesOverlay: overlayActive, nativeSupported, enter, exit, toggle };
}
