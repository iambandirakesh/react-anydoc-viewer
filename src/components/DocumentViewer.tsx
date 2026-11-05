/**
 * The viewer component.
 *
 * One component serves every supported format, because the parser hands it the
 * same model whatever the file was. Responsibilities are split by seam:
 *  - `useDocument` owns loading, cancellation and media lifetime;
 *  - `usePageNavigation`, `useZoom` and `useViewerKeyboard` own page, scale and
 *    input state;
 *  - `SlideCanvas` and `WorkbookView` own the stage for pages and worksheets;
 *  - this component wires them together, owns the chrome (toolbar, thumbnails,
 *    diagnostics) and the public imperative handle.
 */

import {
  forwardRef,
  useCallback,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
/**
 * This import only tells the library build which stylesheet to emit: Vite
 * extracts it to `dist/styles.css` and drops it from the published JS. The
 * component does not load its own styles; consumers import
 * `react-anydoc-viewer/styles.css` themselves.
 */
import "../styles.css";
import type { DocumentFormat } from "../errors";
import type { DocumentModel } from "../model/types";
import { emuToPx } from "../parser/units";
import { useDocument } from "../hooks/useDocument";
import { useElementSize } from "../hooks/useElementSize";
import { useFullscreen } from "../hooks/useFullscreen";
import { usePageNavigation } from "../hooks/usePageNavigation";
import { usePresentationHint } from "../hooks/usePresentationHint";
import { usePresentationMode } from "../hooks/usePresentationMode";
import { useRenderIssues } from "../hooks/useRenderIssues";
import { useViewerKeyboard } from "../hooks/useViewerKeyboard";
import { MAX_ZOOM, MIN_ZOOM, useZoom } from "../hooks/useZoom";
import { SlideCanvas } from "./SlideCanvas";
import { WorkbookView } from "./WorkbookView";
import { ThumbnailSidebar } from "./ThumbnailSidebar";
import { ViewerToolbar } from "./ViewerToolbar";
import { DiagnosticsPanel } from "./DiagnosticsPanel";
import { mergeLabels } from "../labels";
import type { DocumentViewerHandle, DocumentViewerProps } from "../types/public";

const STAGE_INSET = 12;
/**
 * Extra space kept clear at the bottom of the stage while the floating toolbar
 * is on screen, so a fitted slide is never hidden behind it.
 */
const TOOLBAR_CLEARANCE = 44;

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

  /* ---------------------------------------------------------- loading */

  const onLoadRef = useRef(props.onLoad);
  onLoadRef.current = props.onLoad;
  const onErrorRef = useRef(props.onError);
  onErrorRef.current = props.onError;
  const onDiagnosticsRef = useRef(props.onDiagnostics);
  onDiagnosticsRef.current = props.onDiagnostics;

  const handleLoaded = useCallback((parsed: DocumentModel, format: DocumentFormat) => {
    onLoadRef.current?.({
      format,
      pageCount: parsed.slides.length,
      slideCount: parsed.slides.length,
      widthEmu: parsed.width,
      heightEmu: parsed.height,
      aspectRatio: parsed.height > 0 ? parsed.width / parsed.height : 0,
      title: parsed.title,
      diagnostics: parsed.diagnostics,
    });
  }, []);

  const loaded = useDocument(resolvedSource, {
    limits,
    fetchOptions,
    maxDownloadBytes,
    pdfWorkerSrc,
    contentHint,
    onLoad: handleLoaded,
    onError: (error) => onErrorRef.current?.(error),
    onDiagnostics: (diagnostics) => onDiagnosticsRef.current?.(diagnostics),
  });

  const { model, media, status, error, renderPage, content, loadId } = loaded;
  const worksheetPages = content?.kind === "sheet" ? content.pages : undefined;
  const slideCount = model?.slides.length ?? 0;
  const isWorkbook = loaded.format === "xlsx" || loaded.format === "csv";

  /* ------------------------------------------------ navigation, issues */

  const navigation = usePageNavigation({
    count: slideCount,
    controlled: controlledSlide,
    defaultIndex: defaultSlide,
    resetKey: loadId,
    onChange: props.onSlideChange,
  });
  const { index: slideIndex, goTo: goToSlide, next: goNext, previous: goPrevious } = navigation;
  const renderIssues = useRenderIssues(navigation.indexRef, loadId);
  const reportIssue = renderIssues.report;

  /* ------------------------------------------------- presentation mode */

  const [containerRef, size] = useElementSize<HTMLDivElement>();
  const rootRef = useRef<HTMLDivElement | null>(null);
  const fullscreen = useFullscreen(rootRef);

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
  const hintVisible = usePresentationHint(rootRef, presenting);

  /* ------------------------------------------------------ zoom and fit */

  const slideWidth = model ? emuToPx(model.width) : 1280;
  const slideHeight = model ? emuToPx(model.height) : 720;
  // A slideshow fills the viewport: no stage padding, and no toolbar either.
  const inset = presenting ? 0 : STAGE_INSET;
  const zoom = useZoom({
    fit,
    controlled: controlledZoom,
    defaultZoom,
    stage: size,
    pageWidth: slideWidth,
    pageHeight: slideHeight,
    inset,
    bottomInset: presenting || !controlsVisible ? inset : inset + TOOLBAR_CLEARANCE,
    fitDisabled: isWorkbook,
    onZoomChange: props.onZoomChange,
    onFitChange: props.onFitChange,
  });
  const { scale, setZoom, zoomIn, zoomOut, resetZoom } = zoom;

  /* -------------------------------------------------------- thumbnails */

  const [thumbnailsVisible, setThumbnailsVisible] = useState(showThumbnails);
  useEffect(() => setThumbnailsVisible(showThumbnails), [showThumbnails]);
  const [diagnosticsOpen, setDiagnosticsOpen] = useState(debug);
  useEffect(() => setDiagnosticsOpen(debug), [debug]);

  /* ---------------------------------------------------------- keyboard */

  const keyboard = useViewerKeyboard({
    enabled: keyboardNavigation,
    isWorkbook,
    presenting,
    navigation,
    zoomIn,
    zoomOut,
    resetZoom,
    toggleFullscreen: fullscreen.toggle,
    exitPresentation: presentationMode.exit,
  });

  /* ------------------------------------------------------------ handle */

  const scaleRef = useRef(scale);
  scaleRef.current = scale;
  const { countRef: slideCountRef, indexRef: slideIndexRef } = navigation;

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
      reload: loaded.reload,
    }),
    [
      fullscreen,
      goNext,
      goPrevious,
      goToSlide,
      loaded.reload,
      presentationMode,
      resetZoom,
      setZoom,
      slideCountRef,
      slideIndexRef,
      zoomIn,
      zoomOut,
    ],
  );

  /* ---------------------------------------------------------- rendering */

  const effectiveTheme = theme ?? (dark ? "dark" : "light");
  const currentSlide = model && slideCount > 0 ? model.slides[slideIndex] ?? null : null;

  const diagnostics = useMemo(
    () => [...(model?.diagnostics ?? []), ...renderIssues.issues],
    [model, renderIssues.issues],
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
      renderError?.(error, loaded.reload) ?? (
        <div className="radv-status radv-status-error" role="alert" data-testid="rv-error">
          <strong>{labels.errorTitle}</strong>
          <p className="radv-error-message">{error.message}</p>
          <code className="radv-error-code" data-testid="rv-error-code">
            {error.code}
          </code>
          {error.detail ? <p className="radv-error-detail">{error.detail}</p> : null}
          <button type="button" className="radv-btn radv-btn-primary" onClick={loaded.reload}>
            {labels.retry}
          </button>
        </div>
      );
  } else if (status === "ready" && model && media) {
    body = isWorkbook && renderPage && currentSlide ? (
      <WorkbookView
        model={model}
        media={media}
        sheet={currentSlide}
        sheetIndex={slideIndex}
        worksheetPages={worksheetPages}
        renderPage={renderPage}
        showControls={controlsVisible}
        presenting={presenting}
        present={
          showPresentationControl
            ? { label: labels.presentSheet, available: presentationMode.available, enter: presentationMode.enter }
            : undefined
        }
        loadId={loadId}
        rootRef={rootRef}
        viewportRef={containerRef}
        viewportSize={size}
        zoom={zoom}
        labels={labels}
        goToSheet={goToSlide}
        onRenderIssue={reportIssue}
      />
    ) : (
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
      onKeyDown={keyboard.onKeyDown}
      onClick={keyboard.onClick}
      role="region"
      aria-label={labels.viewerRegion}
    >
      <div className="radv-body">
        {!isWorkbook && !presenting && thumbnailsVisible && model && media && status === "ready" ? (
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
          {controlsVisible && !presenting && !isWorkbook ? (
            <ViewerToolbar
              labels={labels}
              slideIndex={slideIndex}
              slideCount={slideCount}
              zoom={scale}
              fit={zoom.fit}
              explicitZoom={zoom.explicit !== null}
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
              onFit={zoom.changeFit}
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
          <span className="radv-present-text">
            {isWorkbook ? labels.sheetPresentationHint : labels.presentationHint}
          </span>
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
