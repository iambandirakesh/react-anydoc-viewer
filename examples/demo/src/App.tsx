/**
 * Demo application.
 *
 * Loads either the built-in sample deck (generated in the browser) or a `.pptx`
 * file dropped onto the page, and exposes the viewer options that matter most
 * while evaluating the library.
 */

import { useCallback, useMemo, useRef, useState, type ChangeEvent, type DragEvent } from "react";
import {
  DocumentViewer,
  detectPdfEngine,
  type FitMode,
  type DocumentViewerErrorType,
  type DocumentViewerHandle,
  type DocumentInfo,
} from "react-anydoc-viewer";
import "react-anydoc-viewer/styles.css";
import { SAMPLE_FILE_NAME, buildSampleDeck } from "./sampleDeck";
import {
  DOCX_MIME,
  SAMPLE_DOCX_FILE_NAME,
  SAMPLE_PDF_FILE_NAME,
  buildSampleDocument,
  buildSamplePdf,
} from "./sampleDocument";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const PDF_MIME = "application/pdf";
/** Everything the viewer opens, for the file picker. */
const ACCEPT = [
  ".pptx",
  ".docx",
  ".pdf",
  PPTX_MIME,
  DOCX_MIME,
  PDF_MIME,
].join(",");

const FORMAT_LABELS: Record<string, string> = {
  pptx: "PowerPoint presentation",
  docx: "Word document",
  pdf: "PDF document",
};

/**
 * pdf.js ships two builds and refuses to pair an API with the other build's
 * worker, so the demo publishes the worker that matches this browser. Both
 * specifiers are written out literally because the bundler has to see them.
 */
const PDF_WORKER_SRC =
  detectPdfEngine() === "legacy"
    ? new URL("pdfjs-dist/legacy/build/pdf.worker.min.mjs", import.meta.url).href
    : new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url).href;

export function App() {
  const [source, setSource] = useState<File | null>(null);
  const [sourceLabel, setSourceLabel] = useState("No file loaded");
  const [showThumbnails, setShowThumbnails] = useState(true);
  const [debug, setDebug] = useState(false);
  const [dark, setDark] = useState(false);
  const [presenting, setPresenting] = useState(false);
  const [fit, setFit] = useState<FitMode>("page");
  const [info, setInfo] = useState<DocumentInfo | null>(null);
  const [error, setError] = useState<DocumentViewerErrorType | null>(null);
  const [slide, setSlide] = useState(0);
  const [dragActive, setDragActive] = useState(false);
  const [buildError, setBuildError] = useState<string | null>(null);
  const viewer = useRef<DocumentViewerHandle>(null);

  const loadBytes = useCallback(
    (bytes: Uint8Array, name: string, type: string) => {
      const file = new File([bytes as BlobPart], name, { type });
      setSource(file);
      setSourceLabel(`${name} (${Math.round(bytes.byteLength / 1024)} KB, generated here)`);
      setBuildError(null);
    },
    [],
  );

  const loadSample = useCallback(() => {
    try {
      loadBytes(buildSampleDeck(), SAMPLE_FILE_NAME, PPTX_MIME);
    } catch (cause) {
      setBuildError(cause instanceof Error ? cause.message : String(cause));
    }
  }, [loadBytes]);

  const loadSampleDocx = useCallback(() => {
    try {
      loadBytes(buildSampleDocument(), SAMPLE_DOCX_FILE_NAME, DOCX_MIME);
    } catch (cause) {
      setBuildError(cause instanceof Error ? cause.message : String(cause));
    }
  }, [loadBytes]);

  const loadSamplePdf = useCallback(() => {
    try {
      loadBytes(buildSamplePdf(), SAMPLE_PDF_FILE_NAME, PDF_MIME);
    } catch (cause) {
      setBuildError(cause instanceof Error ? cause.message : String(cause));
    }
  }, [loadBytes]);

  const accept = useCallback((file: File | undefined) => {
    if (!file) return;
    setSource(file);
    setSourceLabel(`${file.name} (${Math.round(file.size / 1024)} KB)`);
  }, []);

  const onPick = useCallback(
    (event: ChangeEvent<HTMLInputElement>) => {
      accept(event.target.files?.[0]);
    },
    [accept],
  );

  const onDrop = useCallback(
    (event: DragEvent<HTMLDivElement>) => {
      event.preventDefault();
      setDragActive(false);
      accept(event.dataTransfer.files?.[0]);
    },
    [accept],
  );

  const onLoad = useCallback((loaded: DocumentInfo) => {
    setInfo(loaded);
    setError(null);
  }, []);

  const onError = useCallback((viewerError: DocumentViewerErrorType) => {
    setError(viewerError);
    setInfo(null);
  }, []);

  const snippet = useMemo(
    () =>
      `<DocumentViewer
  file={pptxFile}${showThumbnails ? "" : "\n  showThumbnails={false}"}${debug ? "\n  debug" : ""}${
        dark ? "\n  dark" : ""
      }${fit !== "page" ? `\n  fit="${fit}"` : ""}
  onLoad={(info) => …}
  onError={(error) => …}
  onSlideChange={(index) => …}
/>`,
    [dark, debug, fit, showThumbnails],
  );

  return (
    <div className="demo" data-theme={dark ? "dark" : "light"}>
      <header className="demo__header">
        <div>
          <h1>react-anydoc-viewer</h1>
          <p>
            A free, watermark-free React viewer for <code>.pptx</code>, <code>.docx</code> and <code>.pdf</code> files.
            Everything renders from the file in your browser — no server, no API key, no conversion service.
          </p>
        </div>
        <div className="demo__header-actions">
          <button type="button" className="demo__button demo__button--primary" onClick={loadSample}>
            Load sample deck
          </button>
          <button type="button" className="demo__button" onClick={loadSampleDocx}>
            Load sample .docx
          </button>
          <button type="button" className="demo__button" onClick={loadSamplePdf}>
            Load sample .pdf
          </button>
          <label className="demo__button">
            Open a file
            <input type="file" accept={ACCEPT} onChange={onPick} />
          </label>
        </div>
      </header>

      <div className="demo__body">
        <aside className="demo__panel">
          <h2>Options</h2>
          <label className="demo__check">
            <input type="checkbox" checked={showThumbnails} onChange={(e) => setShowThumbnails(e.target.checked)} />
            Thumbnail sidebar
          </label>
          <label className="demo__check">
            <input type="checkbox" checked={debug} onChange={(e) => setDebug(e.target.checked)} />
            Debug mode (diagnostics panel)
          </label>
          <label className="demo__check">
            <input type="checkbox" checked={dark} onChange={(e) => setDark(e.target.checked)} />
            Dark chrome
          </label>
          <label className="demo__check">
            <input
              type="checkbox"
              checked={presenting}
              onChange={(e) => setPresenting(e.target.checked)}
            />
            Presentation mode
          </label>
          <label className="demo__field">
            Fit
            <select value={fit} onChange={(e) => setFit(e.target.value as FitMode)}>
              <option value="page">page</option>
              <option value="width">width</option>
              <option value="actual">actual</option>
              <option value="none">none</option>
            </select>
          </label>

          <h2>Programmatic control</h2>
          <div className="demo__row">
            <button type="button" className="demo__button" onClick={() => viewer.current?.first()}>
              first()
            </button>
            <button type="button" className="demo__button" onClick={() => viewer.current?.previous()}>
              previous()
            </button>
            <button type="button" className="demo__button" onClick={() => viewer.current?.next()}>
              next()
            </button>
            <button type="button" className="demo__button" onClick={() => viewer.current?.last()}>
              last()
            </button>
            <button type="button" className="demo__button" onClick={() => viewer.current?.zoomIn()}>
              zoomIn()
            </button>
            <button type="button" className="demo__button" onClick={() => viewer.current?.resetZoom()}>
              resetZoom()
            </button>
            <button
              type="button"
              className="demo__button"
              onClick={() => viewer.current?.togglePresentation()}
            >
              togglePresentation()
            </button>
          </div>

          <h2>State</h2>
          <dl className="demo__stats">
            <dt>Source</dt>
            <dd>{sourceLabel}</dd>
            <dt>Format</dt>
            <dd>{info ? (FORMAT_LABELS[info.format] ?? info.format) : "—"}</dd>
            <dt>Slides / pages</dt>
            <dd>{info ? info.slideCount : "—"}</dd>
            <dt>Size</dt>
            <dd>{info ? `${(info.widthEmu / 914400).toFixed(2)}in × ${(info.heightEmu / 914400).toFixed(2)}in` : "—"}</dd>
            <dt>Aspect</dt>
            <dd>{info ? info.aspectRatio.toFixed(3) : "—"}</dd>
            <dt>Current slide</dt>
            <dd>{info ? `${slide + 1} / ${info.slideCount}` : "—"}</dd>
            <dt>Diagnostics</dt>
            <dd>{info ? info.diagnostics.length : "—"}</dd>
            <dt>Presenting</dt>
            <dd>{presenting ? "yes" : "no"}</dd>
            <dt>Error</dt>
            <dd>{error ? `${error.code}: ${error.message}` : "none"}</dd>
          </dl>

          <h2>Usage</h2>
          <pre className="demo__code">{snippet}</pre>
          {buildError ? <p className="demo__build-error">Sample deck failed to build: {buildError}</p> : null}
          <p className="demo__hint">
            Every sample is generated in the browser. The deck includes text, shapes, images, a group, a table,
            charts, an empty slide and deliberately unsupported content; the Word document adds headings, lists, a
            merged table, a picture and page breaks; the PDF is rasterised page by page with pdf.js. Turn on debug mode
            to read the diagnostics. Presentation mode hides every control and fills the screen — click or press Space
            to advance, Esc to come back.
          </p>
        </aside>

        <main
          className={`demo__stage${dragActive ? " demo__stage--drag" : ""}`}
          onDragOver={(event) => {
            event.preventDefault();
            setDragActive(true);
          }}
          onDragLeave={() => setDragActive(false)}
          onDrop={onDrop}
        >
          {source ? (
            <DocumentViewer
              ref={viewer}
              file={source}
              /* A published worker keeps PDF parsing off the main thread. */
              pdfWorkerSrc={PDF_WORKER_SRC}
              showThumbnails={showThumbnails}
              allowThumbnailToggle
              debug={debug}
              dark={dark}
              fit={fit}
              presentation={presenting}
              onPresentationChange={setPresenting}
              onLoad={onLoad}
              onError={onError}
              onSlideChange={setSlide}
            />
          ) : (
            <div className="demo__empty">
              <p>Drop a .pptx, .docx or .pdf file here, or</p>
              <button type="button" className="demo__button demo__button--primary" onClick={loadSample}>
                load the sample deck
              </button>
            </div>
          )}
          {dragActive ? <div className="demo__drop">Drop to load</div> : null}
        </main>
      </div>
    </div>
  );
}
