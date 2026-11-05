import { useCallback, useEffect, type KeyboardEvent, type MouseEvent } from "react";
import type { PageNavigation } from "./usePageNavigation";

function isEditableTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  const tag = target.tagName.toLowerCase();
  return tag === "input" || tag === "textarea" || tag === "select" || target.isContentEditable;
}

/** Keys a worksheet leaves to the viewer; everything else moves the cell cursor. */
const WORKBOOK_VIEWER_KEYS = ["+", "=", "-", "_", "0", "f"];

export interface UseViewerKeyboardOptions {
  /** the `keyboardNavigation` prop */
  enabled: boolean;
  /** worksheets keep arrow keys and clicks for themselves */
  isWorkbook: boolean;
  presenting: boolean;
  navigation: Pick<PageNavigation, "goTo" | "next" | "previous" | "countRef">;
  zoomIn: () => void;
  zoomOut: () => void;
  resetZoom: () => void;
  toggleFullscreen: () => void;
  exitPresentation: () => void;
}

export interface ViewerKeyboard {
  /** for the viewer root while it has focus */
  onKeyDown: (event: KeyboardEvent<HTMLDivElement>) => void;
  /** clicking a slide during a slideshow advances it */
  onClick: (event: MouseEvent<HTMLDivElement>) => void;
}

/**
 * Keyboard and click handling.
 *
 * Normally keys are handled on the focused viewer. A slideshow owns the screen
 * instead: keys are handled on the document (focus can sit anywhere after a
 * click) and must not also reach the toolbar handler.
 */
export function useViewerKeyboard({
  enabled,
  isWorkbook,
  presenting,
  navigation,
  zoomIn,
  zoomOut,
  resetZoom,
  toggleFullscreen,
  exitPresentation,
}: UseViewerKeyboardOptions): ViewerKeyboard {
  const { goTo, next, previous, countRef } = navigation;

  const onKeyDown = useCallback(
    (event: KeyboardEvent<HTMLDivElement>) => {
      if (!enabled || event.defaultPrevented) return;
      if (isWorkbook && !WORKBOOK_VIEWER_KEYS.includes(event.key)) return;
      if (event.ctrlKey || event.metaKey || event.altKey) return;
      if (isEditableTarget(event.target)) return;
      let handled = true;
      switch (event.key) {
        case "ArrowRight":
        case "ArrowDown":
        case "PageDown":
        case " ":
        case "n":
          next();
          break;
        case "ArrowLeft":
        case "ArrowUp":
        case "PageUp":
        case "p":
          previous();
          break;
        case "Home":
          goTo(0);
          break;
        case "End":
          goTo(countRef.current - 1);
          break;
        case "+":
        case "=":
          zoomIn();
          break;
        case "-":
        case "_":
          zoomOut();
          break;
        case "0":
          resetZoom();
          break;
        case "f":
          // The overlay fallback means fullscreen always works, even where the
          // native Fullscreen API is unavailable.
          toggleFullscreen();
          break;
        default:
          handled = false;
      }
      if (handled) event.preventDefault();
    },
    [countRef, enabled, goTo, isWorkbook, next, previous, resetZoom, toggleFullscreen, zoomIn, zoomOut],
  );

  useEffect(() => {
    if (!presenting || typeof document === "undefined") return undefined;

    const onPresentingKeyDown = (event: globalThis.KeyboardEvent) => {
      if (event.ctrlKey || event.metaKey || event.altKey) return;
      if (event.key === "Escape") {
        // Esc leaves the slideshow first, even if the fullscreen overlay is on.
        event.preventDefault();
        event.stopImmediatePropagation();
        exitPresentation();
        return;
      }
      if (isWorkbook || !enabled || isEditableTarget(event.target)) return;
      let handled = true;
      switch (event.key) {
        case "ArrowRight":
        case "ArrowDown":
        case "PageDown":
        case "Enter":
        case " ":
        case "n":
          next();
          break;
        case "ArrowLeft":
        case "ArrowUp":
        case "PageUp":
        case "Backspace":
        case "p":
          previous();
          break;
        case "Home":
          goTo(0);
          break;
        case "End":
          goTo(countRef.current - 1);
          break;
        default:
          handled = false;
      }
      if (handled) {
        event.preventDefault();
        event.stopPropagation();
      }
    };

    document.addEventListener("keydown", onPresentingKeyDown, true);
    return () => document.removeEventListener("keydown", onPresentingKeyDown, true);
  }, [countRef, enabled, exitPresentation, goTo, isWorkbook, next, presenting, previous]);

  // Clicking the slide advances, the way PowerPoint does; links keep working.
  const onClick = useCallback(
    (event: MouseEvent<HTMLDivElement>) => {
      if (!presenting || isWorkbook) return;
      const target = event.target as HTMLElement | null;
      if (target?.closest("a,button")) return;
      next();
    },
    [presenting, next, isWorkbook],
  );

  return { onKeyDown, onClick };
}
