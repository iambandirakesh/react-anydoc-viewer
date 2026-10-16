/**
 * The viewer component.
 *
 * One component serves every supported format, because the parser hands it the
 * same model whatever the file was. Responsibilities are split by seam:
 *  - `usePresentation` owns loading, cancellation and media lifetime;
 *  - `SlideCanvas` owns scaling and per-slide failure containment;
 *  - this component owns the chrome (toolbar, thumbnails, diagnostics),
 *    navigation state and the public imperative handle.
 */

import {
  forwardRef,
  useCallback,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent,
  type MouseEvent,
  type ReactNode,
} from "react";
/**
 * The stylesheet is imported here (not from `index.ts`) so that the built
 * declarations never contain a CSS import: consumers would otherwise need a
 * `*.css` module declaration of their own to type-check the package.
 */
import "../styles.css";
import type { DocumentFormat } from "../errors";
import type { PresentationModel } from "../model/types";
import { emuToPx } from "../parser/units";
import { usePresentation } from "../hooks/usePresentation";
import { useElementSize } from "../hooks/useElementSize";
import { useFullscreen } from "../hooks/useFullscreen";
import { usePresentationMode } from "../hooks/usePresentationMode";
import { SlideCanvas } from "./SlideCanvas";
import { ThumbnailSidebar } from "./ThumbnailSidebar";
import { ViewerToolbar } from "./ViewerToolbar";
import { DiagnosticsPanel } from "./DiagnosticsPanel";
import { mergeLabels } from "./labels";
import type {
  FitMode,
  DocumentDiagnostic,
  DocumentViewerHandle,
  DocumentViewerProps,
  DocumentInfo,
} from "../types/public";

const ZOOM_LADDER = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 2, 3, 4];
const STAGE_INSET = 12;
/**
 * Extra space kept clear at the bottom of the stage while the floating toolbar
 * is on screen, so a fitted slide is never hidden behind it.
 */
const TOOLBAR_CLEARANCE = 44;
/** how long the presentation hint stays on screen after the last activity */
const HINT_VISIBLE_MS = 3200;
const MAX_RENDER_ISSUES = 40;
const MIN_ZOOM = 0.1;
const MAX_ZOOM = 8;

function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value;
}

function isEditableTarget(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) return false;
  const tag = target.tagName.toLowerCase();
  return tag === "input" || tag === "textarea" || tag === "select" || target.isContentEditable;
}

export const DocumentViewer = forwardRef<DocumentViewerHandle, DocumentViewerProps>(function DocumentViewer(
  props,
  ref,
): ReactNode {
  const {
    source,
    file,
    slide: controlledSlide,
    defaultSlide = 0,
    showControls = true,
    showToolbar,
    showThumbnails = false,
    allowThumbnailToggle,
    fit = "page",
    zoom: controlledZoom,
    defaultZoom,
    presentation: controlledPresentation,
    defaultPresentation = false,
    showPresentationControl = true,
    keyboardNavigation = true,
    theme,
    dark = false,
    className,
    style,
    emptySlideMessage,
    labels: labelOverrides,
    debug = false,
    limits,
    fetchOptions,
    maxDownloadBytes,
    pdfWorkerSrc,
    contentHint,
    renderSlideError,
    renderLoader,
    renderError,
  } = props;

  const labels = useMemo(() => mergeLabels(labelOverrides), [labelOverrides]);
  const resolvedSource = source ?? file ?? null;
  const controlsVisible = showToolbar ?? showControls;

  /* ------------------------------------------------------ render issues */

  const [renderIssues, setRenderIssues] = useState<DocumentDiagnostic[]>([]);
  const issueKeys = useRef<Set<string>>(new Set());
  const pendingIssues = useRef<DocumentDiagnostic[]>([]);
  const slideIndexRef = useRef(0);

  /**
   * Issues are reported from inside the renderer, i.e. while a child component is
   * rendering. Updating state there trips React's "cannot update a component while
   * rendering a different component" warning (and concurrent rendering may drop
   * the update), so reports are buffered and flushed right after the commit.
   */
  const reportIssue = useCallback((code: string, message: string) => {
    const slideIndex = slideIndexRef.current;
    const key = `${slideIndex}|${code}|${message}`;
    if (issueKeys.current.has(key) || issueKeys.current.size >= MAX_RENDER_ISSUES) return;
    issueKeys.current.add(key);
    pendingIssues.current.push({ severity: "warning", code, message, slideIndex });
  }, []);

  useEffect(() => {
    if (pendingIssues.current.length === 0) return;
    const queued = pendingIssues.current;
    pendingIssues.current = [];
    setRenderIssues((previous) => [...previous, ...queued]);
  });

  /* ---------------------------------------------------------- loading */

  const onLoadRef = useRef(props.onLoad);
  onLoadRef.current = props.onLoad;
  const onErrorRef = useRef(props.onError);
  onErrorRef.current = props.onError;
  const onDiagnosticsRef = useRef(props.onDiagnostics);
  onDiagnosticsRef.current = props.onDiagnostics;

  const handleLoaded = useCallback((parsed: PresentationModel, format: DocumentFormat) => {
    onLoadRef.current?.({
      format,
      slideCount: parsed.slides.length,
      widthEmu: parsed.width,
      heightEmu: parsed.height,
      aspectRatio: parsed.height > 0 ? parsed.width / parsed.height : 0,
      title: parsed.title,
      diagnostics: parsed.diagnostics,
    });
  }, []);

  const presentation = usePresentation(resolvedSource, {
    limits,
    fetchOptions,
    maxDownloadBytes,
    pdfWorkerSrc,
    contentHint,
    onLoad: handleLoaded,
    onError: (error) => onErrorRef.current?.(error),
    onDiagnostics: (diagnostics) => onDiagnosticsRef.current?.(diagnostics),
  });

  const { model, media, status, error, renderPage } = presentation;
  const slideCount = model?.slides.length ?? 0;

  /* ------------------------------------------------------- navigation */

  const [internalSlide, setInternalSlide] = useState(defaultSlide);
  const requestedIndex = controlledSlide ?? internalSlide;
  const slideIndex = slideCount === 0 ? 0 : clamp(Math.round(requestedIndex), 0, slideCount - 1);
  slideIndexRef.current = slideIndex;

  const onSlideChangeRef = useRef(props.onSlideChange);
  onSlideChangeRef.current = props.onSlideChange;

  const goToSlide = useCallback(
    (index: number) => {
      const currentCount = slideCountRef.current;
      if (currentCount === 0) return;
      const next = clamp(Math.round(index), 0, currentCount - 1);
      if (controlledSlideRef.current === undefined) setInternalSlide(next);
      if (next !== slideIndexRef.current) onSlideChangeRef.current?.(next);
    },
    [],
  );

  const slideCountRef = useRef(slideCount);
  slideCountRef.current = slideCount;
  const controlledSlideRef = useRef(controlledSlide);
  controlledSlideRef.current = controlledSlide;

  const goNext = useCallback(() => goToSlide(slideIndexRef.current + 1), [goToSlide]);
  const goPrevious = useCallback(() => goToSlide(slideIndexRef.current - 1), [goToSlide]);

  // Keep a controlled/out-of-range index inside the new presentation's bounds.
  useEffect(() => {
    if (slideCount === 0) return;
    if (requestedIndex < 0 || requestedIndex > slideCount - 1 || !Number.isFinite(requestedIndex)) {
      setInternalSlide(slideIndex);
      onSlideChangeRef.current?.(slideIndex);
    }
  }, [requestedIndex, slideCount, slideIndex]);

  // A newly loaded presentation resets to the first slide and clears issues.
  const loadId = presentation.loadId;
  useEffect(() => {
    if (loadId === 0) return;
    issueKeys.current = new Set();
    pendingIssues.current = [];
    setRenderIssues([]);
    setInternalSlide(defaultSlide);
  }, [loadId, defaultSlide]);

  /* ------------------------------------------------------ zoom and fit */

  const [containerRef, size] = useElementSize<HTMLDivElement>();
  const rootRef = useRef<HTMLDivElement | null>(null);
  const fullscreen = useFullscreen(rootRef);

  /* ------------------------------------------------- presentation mode */

  const onPresentationChangeRef = useRef(props.onPresentationChange);
  onPresentationChangeRef.current = props.onPresentationChange;
  const presentationMode = usePresentationMode({
    controlled: controlledPresentation,
    initial: defaultPresentation,
    available: status === "ready" && slideCount > 0,
    onActiveChange: (active) => onPresentationChangeRef.current?.(active),
  });
  // A slideshow only takes over the screen when there is something to show; an
  // error or an empty source still renders its own state (and the state stays
  // controlled, so a host app can decide what to do about it).
  const presenting = presentationMode.active && presentationMode.available;

  // A slide counter and key hint that fades out; any pointer movement brings it back.
  const [hintVisible, setHintVisible] = useState(false);
  useEffect(() => {
    if (!presenting) {
      setHintVisible(false);
      return undefined;
    }
    setHintVisible(true);
    const root = rootRef.current;
    let timer: ReturnType<typeof setTimeout> | null = null;
    const show = () => {
      setHintVisible(true);
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => setHintVisible(false), HINT_VISIBLE_MS);
    };
    show();
    root?.addEventListener("mousemove", show);
    root?.addEventListener("touchstart", show);
    return () => {
      if (timer) clearTimeout(timer);
      root?.removeEventListener("mousemove", show);
      root?.removeEventListener("touchstart", show);
    };
  }, [presenting]);

  const [fitState, setFitState] = useState<FitMode>(fit);
  useEffect(() => setFitState(fit), [fit]);

  const [internalZoom, setInternalZoom] = useState<number | null>(defaultZoom ?? null);
  useEffect(() => {
    if (controlledZoom !== undefined) setInternalZoom(controlledZoom);
  }, [controlledZoom]);

  const slideWidth = model ? emuToPx(model.width) : 1280;
  const slideHeight = model ? emuToPx(model.height) : 720;

  const fitScale = useMemo(() => {
    if (!size.measured) return 1;
    // A slideshow fills the viewport: no stage padding, and no toolbar either.
    const inset = presenting ? 0 : STAGE_INSET;
    const bottomInset = presenting || !controlsVisible ? inset : inset + TOOLBAR_CLEARANCE;
    const availableWidth = Math.max(size.width - inset * 2, 1);
    const availableHeight = Math.max(size.height - inset - bottomInset, 1);
    switch (fitState) {
      case "width":
        return availableWidth / slideWidth;
      case "height":
        return availableHeight / slideHeight;
      case "page":
        return Math.min(availableWidth / slideWidth, availableHeight / slideHeight);
      default:
        return 1;
    }
  }, [
    controlsVisible,
    fitState,
    presenting,
    size.height,
    size.measured,
    size.width,
    slideHeight,
    slideWidth,
  ]);

  const scale = internalZoom ?? fitScale;
  const onZoomChangeRef = useRef(props.onZoomChange);
  onZoomChangeRef.current = props.onZoomChange;
  const onFitChangeRef = useRef(props.onFitChange);
  onFitChangeRef.current = props.onFitChange;

  const setZoom = useCallback(
    (next: number) => {
      const clamped = clamp(next, MIN_ZOOM, MAX_ZOOM);
      if (controlledZoom === undefined) setInternalZoom(clamped);
      onZoomChangeRef.current?.(clamped);
    },
    [controlledZoom],
  );

  const zoomIn = useCallback(() => {
    const next = ZOOM_LADDER.find((step) => step > scale + 0.001) ?? ZOOM_LADDER[ZOOM_LADDER.length - 1]!;
    setZoom(next);
  }, [scale, setZoom]);

  const zoomOut = useCallback(() => {
    const previous = [...ZOOM_LADDER].reverse().find((step) => step < scale - 0.001) ?? ZOOM_LADDER[0]!;
    setZoom(previous);
  }, [scale, setZoom]);

  const resetZoom = useCallback(() => {
    if (controlledZoom === undefined) setInternalZoom(null);
    onZoomChangeRef.current?.(fitScale);
  }, [controlledZoom, fitScale]);

  const changeFit = useCallback(
    (next: FitMode) => {
      setFitState(next);
      if (controlledZoom === undefined) setInternalZoom(null);
      onFitChangeRef.current?.(next);
    },
    [controlledZoom],
  );

  /* -------------------------------------------------------- thumbnails */

  const [thumbnailsVisible, setThumbnailsVisible] = useState(showThumbnails);
  useEffect(() => setThumbnailsVisible(showThumbnails), [showThumbnails]);
  const [diagnosticsOpen, setDiagnosticsOpen] = useState(debug);
  useEffect(() => setDiagnosticsOpen(debug), [debug]);

  /* ---------------------------------------------------------- keyboard */

  const onKeyDown = useCallback(
    (event: KeyboardEvent<HTMLDivElement>) => {
      if (!keyboardNavigation || event.defaultPrevented) return;
      if (event.ctrlKey || event.metaKey || event.altKey) return;
      if (isEditableTarget(event.target)) return;
      let handled = true;
      switch (event.key) {
        case "ArrowRight":
        case "ArrowDown":
        case "PageDown":
        case " ":
        case "n":
          goNext();
          break;
        case "ArrowLeft":
        case "ArrowUp":
        case "PageUp":
        case "p":
          goPrevious();
          break;
        case "Home":
          goToSlide(0);
          break;
        case "End":
          goToSlide(slideCountRef.current - 1);
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
          fullscreen.toggle();
          break;
        default:
          handled = false;
      }
      if (handled) event.preventDefault();
    },
    [fullscreen, goNext, goPrevious, goToSlide, keyboardNavigation, resetZoom, zoomIn, zoomOut],
  );

  // A slideshow owns the screen: keys are handled on the document (focus can sit
  // anywhere after a click) and must not also reach the toolbar handler.
  useEffect(() => {
    if (!presenting || typeof document === "undefined") return undefined;

    const onPresentingKeyDown = (event: globalThis.KeyboardEvent) => {
      if (event.ctrlKey || event.metaKey || event.altKey) return;
      if (event.key === "Escape") {
        // Esc leaves the slideshow first, even if the fullscreen overlay is on.
        event.preventDefault();
        event.stopImmediatePropagation();
        presentationMode.exit();
        return;
      }
      if (!keyboardNavigation || isEditableTarget(event.target)) return;
      let handled = true;
      switch (event.key) {
        case "ArrowRight":
        case "ArrowDown":
        case "PageDown":
        case "Enter":
        case " ":
        case "n":
          goNext();
          break;
        case "ArrowLeft":
        case "ArrowUp":
        case "PageUp":
        case "Backspace":
        case "p":
          goPrevious();
          break;
        case "Home":
          goToSlide(0);
          break;
        case "End":
          goToSlide(slideCountRef.current - 1);
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
  }, [presenting, keyboardNavigation, goNext, goPrevious, goToSlide, presentationMode]);

  // Clicking the slide advances, the way PowerPoint does; links keep working.
  const onViewerClick = useCallback(
    (event: MouseEvent<HTMLDivElement>) => {
      if (!presenting) return;
      const target = event.target as HTMLElement | null;
      if (target?.closest("a,button")) return;
      goNext();
    },
    [presenting, goNext],
  );

  /* ------------------------------------------------------------ handle */

  const scaleRef = useRef(scale);
  scaleRef.current = scale;

  useImperativeHandle(
    ref,
    (): DocumentViewerHandle => ({
      goToSlide,
      next: goNext,
      previous: goPrevious,
      first: () => goToSlide(0),
      last: () => goToSlide(slideCountRef.current - 1),
      getCurrentSlide: () => slideIndexRef.current,
      getSlideCount: () => slideCountRef.current,
      setZoom,
      zoomIn,
      zoomOut,
      resetZoom,
      getZoom: () => scaleRef.current,
      enterFullscreen: fullscreen.enter,
      exitFullscreen: fullscreen.exit,
      toggleFullscreen: fullscreen.toggle,
      enterPresentation: presentationMode.enter,
      exitPresentation: presentationMode.exit,
      togglePresentation: presentationMode.toggle,
      isPresenting: () => presentationMode.active,
      reload: presentation.reload,
    }),
    [
      fullscreen,
      goNext,
      goPrevious,
      goToSlide,
      presentation.reload,
      presentationMode,
      resetZoom,
      setZoom,
      zoomIn,
      zoomOut,
    ],
  );

  /* ---------------------------------------------------------- rendering */

  const effectiveTheme = theme ?? (dark ? "dark" : "light");
  const currentSlide = model && slideCount > 0 ? model.slides[slideIndex] ?? null : null;

  const diagnostics = useMemo(
    () => [...(model?.diagnostics ?? []), ...renderIssues],
    [model, renderIssues],
  );

  let body: ReactNode;
  if (status === "loading") {
    body =
      renderLoader?.() ?? (
        <div className="radv-status" role="status" data-testid="rv-loading">
          <span className="radv-spinner" aria-hidden="true" />
          <span>{labels.loading}</span>
        </div>
      );
  } else if (status === "error" && error) {
    body =
      renderError?.(error, presentation.reload) ?? (
        <div className="radv-status radv-status-error" role="alert" data-testid="rv-error">
          <strong>{labels.errorTitle}</strong>
          <p className="radv-error-message">{error.message}</p>
          <code className="radv-error-code" data-testid="rv-error-code">
            {error.code}
          </code>
          {error.detail ? <p className="radv-error-detail">{error.detail}</p> : null}
          <button type="button" className="radv-btn radv-btn-primary" onClick={presentation.reload}>
            {labels.retry}
          </button>
        </div>
      );
  } else if (status === "ready" && model && media) {
    body = (
      <SlideCanvas
        model={model}
        media={media}
        slide={currentSlide}
        slideIndex={slideIndex}
        slideWidth={slideWidth}
        slideHeight={slideHeight}
        scale={scale}
        size={size}
        containerRef={containerRef}
        labels={labels}
        emptySlideMessage={emptySlideMessage}
        renderSlideError={renderSlideError}
        onRenderIssue={reportIssue}
        renderPage={renderPage}
        bottomClearance={!presenting && controlsVisible ? TOOLBAR_CLEARANCE : 0}
      />
    );
  } else {
    body = <div className="radv-status" role="status" aria-hidden="true" />;
  }

  const overlayClass = fullscreen.usesOverlay ? " radv-root-overlay" : "";
  const presentingClass = presenting ? " radv-root-presenting" : "";

  return (
    <div
      className={`radv-root radv-theme-${effectiveTheme}${overlayClass}${presentingClass}${
        className ? ` ${className}` : ""
      }`}
      style={style}
      ref={rootRef}
      data-testid="rv-root"
      data-status={status}
      data-presenting={presenting ? "true" : "false"}
      data-fullscreen={fullscreen.usesOverlay ? "overlay" : fullscreen.isFullscreen ? "native" : "false"}
      tabIndex={keyboardNavigation ? 0 : -1}
      onKeyDown={onKeyDown}
      onClick={onViewerClick}
      role="region"
      aria-label={labels.viewerRegion}
    >
      <div className="radv-body">
        {!presenting && thumbnailsVisible && model && media && status === "ready" ? (
          <ThumbnailSidebar
            model={model}
            media={media}
            slideWidth={slideWidth}
            slideHeight={slideHeight}
            currentIndex={slideIndex}
            labels={labels}
            onSelect={goToSlide}
            renderPage={renderPage}
          />
        ) : null}
        <div className="radv-stage">
          {body}
          {diagnosticsOpen && !presenting ? (
            <DiagnosticsPanel
              diagnostics={diagnostics}
              labels={labels}
              onClose={() => setDiagnosticsOpen(false)}
            />
          ) : null}
          {/*
           * The toolbar floats over the stage (centred at the bottom) instead of
           * taking a strip off the top, so the slide gets the full width. It is
           * rendered last so it paints above the diagnostics panel, and the
           * stylesheet keeps it out of sight until the viewer is hovered or one of
           * its controls has focus.
           */}
          {controlsVisible && !presenting ? (
            <ViewerToolbar
              labels={labels}
              slideIndex={slideIndex}
              slideCount={slideCount}
              zoom={scale}
              fit={fitState}
              explicitZoom={internalZoom !== null}
              thumbnailsVisible={thumbnailsVisible}
              allowThumbnailToggle={allowThumbnailToggle ?? showThumbnails}
              isFullscreen={fullscreen.isFullscreen}
              canZoomIn={scale < MAX_ZOOM}
              canZoomOut={scale > MIN_ZOOM}
              onPrevious={goPrevious}
              onNext={goNext}
              onZoom={setZoom}
              onZoomIn={zoomIn}
              onZoomOut={zoomOut}
              onFit={changeFit}
              onToggleThumbnails={() => setThumbnailsVisible((visible) => !visible)}
              onToggleFullscreen={fullscreen.toggle}
              showPresentationControl={showPresentationControl}
              presentationAvailable={presentationMode.available}
              onPresent={presentationMode.enter}
            />
          ) : null}
        </div>
      </div>
      {presenting ? (
        <div
          className={`radv-present-hint${hintVisible ? " radv-present-hint-visible" : ""}`}
          data-testid="rv-present-hint"
        >
          <span className="radv-present-text">{labels.presentationHint}</span>
          <span className="radv-present-counter">
            {labels.slideOf(slideCount === 0 ? 0 : slideIndex + 1, slideCount)}
          </span>
          <button
            type="button"
            /* `radv-present-exit` is load-bearing: the pill sits on black, so its
               button deliberately overrides the theme's hover colour, and the
               class is what lets those rules win the cascade. */
            className="radv-btn radv-btn-small radv-present-exit"
            onClick={(event) => {
              event.stopPropagation();
              presentationMode.exit();
            }}
          >
            {labels.exitPresentation}
          </button>
        </div>
      ) : null}
    </div>
  );
});

/**
 * @deprecated The package was called `react-pptx-viewer` before it also rendered
 * `.docx` and `.pdf`; use {@link DocumentViewer}. It is the same component, so
 * props, refs and state keep working unchanged.
 */
export const PPTViewer = DocumentViewer;

export default DocumentViewer;
