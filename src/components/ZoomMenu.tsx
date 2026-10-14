import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type FocusEvent,
  type KeyboardEvent,
  type ReactNode,
} from "react";
import type { FitMode, ViewerLabels } from "../types/public";

/**
 * The zoom stops offered by the dropdown. They are the conventional percentages
 * every slide viewer lists, so a reader can always get back to a familiar scale.
 */
export const ZOOM_STOPS = [0.1, 0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 3, 4, 5];

/** Fit modes the dropdown can switch to, in the order PowerPoint lists them. */
const FIT_STOPS: FitMode[] = ["width", "height", "page", "actual"];

export interface ZoomMenuProps {
  labels: ViewerLabels;
  /** effective scale; `1` is 100% */
  zoom: number;
  /** the fit mode in effect, used to mark the chosen item */
  fit: FitMode;
  /** true when the scale was chosen by hand instead of by a fit mode */
  explicitZoom: boolean;
  onZoom: (scale: number) => void;
  onFit: (fit: FitMode) => void;
}

function itemsOf(menu: HTMLDivElement | null): HTMLButtonElement[] {
  if (!menu) return [];
  return Array.from(menu.querySelectorAll<HTMLButtonElement>('[role="menuitemradio"]'));
}

function sameScale(a: number, b: number): boolean {
  return Math.abs(a - b) < 0.005;
}

/**
 * The zoom percentage in the toolbar, with the zoom stops and the fit modes in a
 * dropdown above it.
 *
 * The trigger is a split control — the percentage, a divider, then a caret — but
 * a single button, so it keeps one accessible name and one tab stop. It opens on
 * click and on `Enter`, `Space`, `ArrowUp` and `ArrowDown`; opening from a click
 * leaves focus on the trigger, while opening from the keyboard moves focus to the
 * current choice, the standard `menuitemradio` behaviour. It never opens on hover:
 * a menu that appears from the pointer passing by covers the slide underneath and
 * is easy to open by accident while reading.
 */
export function ZoomMenu(props: ZoomMenuProps): ReactNode {
  const { labels, zoom, fit, explicitZoom, onZoom, onFit } = props;
  const [open, setOpen] = useState(false);

  const wrapperRef = useRef<HTMLDivElement | null>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const menuRef = useRef<HTMLDivElement | null>(null);
  /** where focus should land the next time the menu opens from the keyboard */
  const pendingFocus = useRef<1 | -1 | null>(null);

  const close = useCallback((restoreFocus: boolean) => {
    setOpen(false);
    if (restoreFocus) triggerRef.current?.focus();
  }, []);

  // A click anywhere else dismisses the menu, the way a popover behaves.
  useEffect(() => {
    if (!open) return undefined;
    const onPointerDown = (event: Event) => {
      if (wrapperRef.current?.contains(event.target as Node)) return;
      setOpen(false);
    };
    document.addEventListener("pointerdown", onPointerDown);
    return () => document.removeEventListener("pointerdown", onPointerDown);
  }, [open]);

  // Keyboard opening focuses the choice in effect, so the menu announces itself.
  useEffect(() => {
    if (!open || !pendingFocus.current) return;
    const items = itemsOf(menuRef.current);
    const checked = items.findIndex((item) => item.getAttribute("aria-checked") === "true");
    const index = checked >= 0 ? checked : pendingFocus.current === 1 ? 0 : items.length - 1;
    items[index]?.focus();
    pendingFocus.current = null;
  }, [open]);

  const openFromKeyboard = useCallback((direction: 1 | -1) => {
    pendingFocus.current = direction;
    setOpen(true);
  }, []);

  const onTriggerKeyDown = useCallback(
    (event: KeyboardEvent<HTMLButtonElement>) => {
      switch (event.key) {
        case "ArrowDown":
        case "ArrowUp":
          event.preventDefault();
          event.stopPropagation();
          openFromKeyboard(event.key === "ArrowDown" ? 1 : -1);
          break;
        case "Escape":
          if (!open) return;
          event.preventDefault();
          event.stopPropagation();
          close(true);
          break;
        default:
          break;
      }
    },
    [close, open, openFromKeyboard],
  );

  const onMenuKeyDown = useCallback(
    (event: KeyboardEvent<HTMLDivElement>) => {
      // The viewer handles Space/arrows globally; inside an open menu they belong
      // to the menu only.
      const items = itemsOf(menuRef.current);
      const current = items.findIndex((item) => item === document.activeElement);
      let next: number | null = null;
      switch (event.key) {
        case "ArrowDown":
          next = current < 0 ? 0 : (current + 1) % items.length;
          break;
        case "ArrowUp":
          next = current <= 0 ? items.length - 1 : current - 1;
          break;
        case "Home":
          next = 0;
          break;
        case "End":
          next = items.length - 1;
          break;
        case "Escape":
          event.preventDefault();
          event.stopPropagation();
          close(true);
          return;
        case "Tab":
          event.stopPropagation();
          setOpen(false);
          return;
        default:
          event.stopPropagation();
          return;
      }
      event.preventDefault();
      event.stopPropagation();
      items[next]?.focus();
    },
    [close],
  );

  const choose = useCallback(
    (apply: () => void) => {
      apply();
      close(true);
    },
    [close],
  );

  const onBlur = useCallback(
    (event: FocusEvent<HTMLDivElement>) => {
      const next = event.relatedTarget as Node | null;
      if (next && wrapperRef.current?.contains(next)) return;
      close(false);
    },
    [close],
  );

  return (
    <div className="radv-zoom-menu" ref={wrapperRef} onBlur={onBlur}>
      <button
        type="button"
        ref={triggerRef}
        className={`radv-btn radv-zoom-trigger${open ? " radv-btn-active" : ""}`}
        aria-haspopup="menu"
        aria-expanded={open}
        aria-label={labels.zoomMenu}
        title={labels.zoomMenu}
        onClick={() => {
          if (open) {
            close(true);
            return;
          }
          pendingFocus.current = null;
          setOpen(true);
        }}
        onKeyDown={onTriggerKeyDown}
      >
        <span className="radv-zoom" aria-live="polite">
          {Math.round(zoom * 100)}%
        </span>
        <span className="radv-caret" aria-hidden="true">
          ▾
        </span>
      </button>

      {open ? (
        <div
          className="radv-menu"
          role="menu"
          aria-label={labels.zoomMenu}
          ref={menuRef}
          onKeyDown={onMenuKeyDown}
        >
          {ZOOM_STOPS.map((stop) => {
            const checked = explicitZoom && sameScale(zoom, stop);
            return (
              <button
                type="button"
                key={`zoom-${stop}`}
                role="menuitemradio"
                aria-checked={checked}
                className="radv-menu-item"
                onClick={() => choose(() => onZoom(stop))}
              >
                <span>{Math.round(stop * 100)}%</span>
                <span className="radv-menu-check" aria-hidden="true">
                  {checked ? "✓" : ""}
                </span>
              </button>
            );
          })}
          <div className="radv-menu-separator" role="separator" />
          {FIT_STOPS.map((mode) => {
            const checked = !explicitZoom && fit === mode;
            return (
              <button
                type="button"
                key={`fit-${mode}`}
                role="menuitemradio"
                aria-checked={checked}
                className="radv-menu-item"
                onClick={() => choose(() => onFit(mode))}
              >
                <span>{fitLabel(mode, labels)}</span>
                <span className="radv-menu-check" aria-hidden="true">
                  {checked ? "✓" : ""}
                </span>
              </button>
            );
          })}
        </div>
      ) : null}
    </div>
  );
}

function fitLabel(mode: FitMode, labels: ViewerLabels): string {
  switch (mode) {
    case "width":
      return labels.fitWidth;
    case "height":
      return labels.fitHeight;
    case "page":
      return labels.fitPage;
    case "actual":
      return labels.actualSize;
    default:
      return labels.fitPage;
  }
}
