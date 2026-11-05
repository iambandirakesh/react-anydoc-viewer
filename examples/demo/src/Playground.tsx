/**
 * The live playground: sample tabs, the viewer, and an inspector with the options,
 * the imperative API and the state the viewer reports.
 */

import { useCallback, useEffect, useMemo, useRef, useState, type ChangeEvent, type DragEvent } from "react";
import {
  ACCEPT_ATTRIBUTE,
  DocumentViewer,
  detectPdfEngine,
  type FitMode,
  type DocumentViewerErrorType,
  type DocumentViewerHandle,
  type DocumentInfo,
  version,
} from "react-anydoc-viewer";
import "react-anydoc-viewer/styles.css";
import { SAMPLES, type Sample } from "./samples";
import { Code } from "./Code";
import { issueUrl } from "./report";

const FORMAT_LABELS: Record<string, string> = {
  pptx: "PowerPoint presentation",
  docx: "Word document",
  xlsx: "Excel workbook",
  pdf: "PDF document",
  image: "image",
  csv: "delimited text table",
  text: "text file",
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

/** `?sample=docx` opens that sample; `?sample=none` starts empty (the benchmarks use it). */
function initialSampleId(): string | null {
  if (typeof window === "undefined") return SAMPLES[0]!.id;
  const requested = new URLSearchParams(window.location.search).get("sample");
  if (requested === "none") return null;
  return SAMPLES.some((sample) => sample.id === requested) ? requested : SAMPLES[0]!.id;
}

export function Playground({ siteDark }: { siteDark: boolean }) {
  const [source, setSource] = useState<File | null>(null);
  const [activeSample, setActiveSample] = useState<string | null>(null);
  const [sourceLabel, setSourceLabel] = useState("No file loaded");
  // A phone has no room for the sidebar next to the page.
  const [showThumbnails, setShowThumbnails] = useState(
    () => typeof window === "undefined" || !window.matchMedia?.("(max-width: 760px)").matches,
  );
  const [debug, setDebug] = useState(false);
  const [dark, setDark] = useState(siteDark);
  const [presenting, setPresenting] = useState(false);
  const [fit, setFit] = useState<FitMode>("page");
  const [info, setInfo] = useState<DocumentInfo | null>(null);
  const [error, setError] = useState<DocumentViewerErrorType | null>(null);
  const [slide, setSlide] = useState(0);
  const [dragActive, setDragActive] = useState(false);
  const [buildError, setBuildError] = useState<string | null>(null);
  const viewer = useRef<DocumentViewerHandle>(null);

  // The viewer's chrome follows the site theme; the checkbox can still override it.
  useEffect(() => setDark(siteDark), [siteDark]);

  const loadSample = useCallback((sample: Sample) => {
    try {
      const bytes = sample.build();
      setSource(new File([bytes as BlobPart], sample.fileName, { type: sample.mime }));
      setActiveSample(sample.id);
      setSourceLabel(`${sample.fileName} (${Math.round(bytes.byteLength / 1024)} KB, generated here)`);
      setBuildError(null);
    } catch (cause) {
      setBuildError(`${sample.fileName} failed to build: ${cause instanceof Error ? cause.message : String(cause)}`);
    }
  }, []);

  useEffect(() => {
    const id = initialSampleId();
    const sample = SAMPLES.find((candidate) => candidate.id === id);
    if (sample) loadSample(sample);
  }, [loadSample]);

  const accept = useCallback((file: File | undefined) => {
    if (!file) return;
    setSource(file);
    setActiveSample(null);
    setSourceLabel(`${file.name} (${Math.round(file.size / 1024)} KB)`);
  }, []);

  const onPick = useCallback((event: ChangeEvent<HTMLInputElement>) => accept(event.target.files?.[0]), [accept]);

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
  file={file}${showThumbnails ? "" : "\n  showThumbnails={false}"}${debug ? "\n  debug" : ""}${dark ? "\n  dark" : ""}${
        fit !== "page" ? `\n  fit="${fit}"` : ""
      }
  onLoad={(info) => …}
  onError={(error) => …}
  onSlideChange={(index) => …}
/>`,
    [dark, debug, fit, showThumbnails],
  );

  const current = SAMPLES.find((sample) => sample.id === activeSample);
  const fileName = source?.name ?? "no file";
  const extension = !current && source?.name.includes(".") ? source.name.slice(source.name.lastIndexOf(".")).toLowerCase() : null;
  const reportHref = issueUrl({ version, info, error, sample: current?.label ?? null, extension });

  return (
    <div
      className={`pg${dragActive ? " pg--drag" : ""}`}
      onDragOver={(event) => {
        event.preventDefault();
        setDragActive(true);
      }}
      onDragLeave={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setDragActive(false);
      }}
      onDrop={onDrop}
    >
      <div className="pg__bar">
        <span className="pg__dots" aria-hidden="true">
          <i />
          <i />
          <i />
        </span>
        <div className="pg__tabs" role="group" aria-label="Sample documents">
          {SAMPLES.map((sample) => (
            <button
              key={sample.id}
              type="button"
              className="pg__tab"
              aria-pressed={activeSample === sample.id}
              onClick={() => loadSample(sample)}
            >
              {sample.label} <span className="pg__ext">{sample.ext}</span>
            </button>
          ))}
        </div>
        <label className="pg__open">
          <svg viewBox="0 0 20 20" aria-hidden="true">
            <path d="M10 3v10M5.5 7.5 10 3l4.5 4.5M4 13v3h12v-3" />
          </svg>
          Open a file
          {/* The picker's filter comes from the library, so it can never disagree
              with what the viewer actually opens. */}
          <input type="file" accept={ACCEPT_ATTRIBUTE} onChange={onPick} />
        </label>
      </div>

      <div className="pg__body">
        <div className="pg__stage">
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
            <div className="pg__empty">
              <svg viewBox="0 0 48 48" aria-hidden="true">
                <path d="M14 6h14l10 10v26H14z" />
                <path d="M28 6v10h10M21 27h10M26 22v10" />
              </svg>
              <p>Drop a .pptx, .docx, .xlsx, .pdf, image, .csv or text file here, or pick a sample above.</p>
            </div>
          )}
          {dragActive ? <div className="pg__drop">Drop to open it here — nothing is uploaded</div> : null}
        </div>

        <aside className="pg__inspector" aria-label="Inspector">
          <section>
            <h3>Options</h3>
            <label className="pg__check">
              <input type="checkbox" checked={showThumbnails} onChange={(e) => setShowThumbnails(e.target.checked)} />
              Thumbnail sidebar
            </label>
            <label className="pg__check">
              <input type="checkbox" checked={debug} onChange={(e) => setDebug(e.target.checked)} />
              Debug mode (diagnostics panel)
            </label>
            <label className="pg__check">
              <input type="checkbox" checked={dark} onChange={(e) => setDark(e.target.checked)} />
              Dark chrome
            </label>
            <label className="pg__check">
              <input type="checkbox" checked={presenting} onChange={(e) => setPresenting(e.target.checked)} />
              Presentation mode
            </label>
            <label className="pg__field">
              Fit
              <select value={fit} onChange={(e) => setFit(e.target.value as FitMode)}>
                <option value="page">page</option>
                <option value="width">width</option>
                <option value="height">height</option>
                <option value="actual">actual</option>
                <option value="none">none</option>
              </select>
            </label>
          </section>

          <section>
            <h3>Imperative API</h3>
            <div className="pg__api">
              <button type="button" onClick={() => viewer.current?.first()}>first()</button>
              <button type="button" onClick={() => viewer.current?.previous()}>previous()</button>
              <button type="button" onClick={() => viewer.current?.next()}>next()</button>
              <button type="button" onClick={() => viewer.current?.last()}>last()</button>
              <button type="button" onClick={() => viewer.current?.zoomIn()}>zoomIn()</button>
              <button type="button" onClick={() => viewer.current?.resetZoom()}>resetZoom()</button>
              <button type="button" onClick={() => viewer.current?.togglePresentation()}>togglePresentation()</button>
            </div>
          </section>

          <section>
            <h3>State</h3>
            <dl className="demo__stats">
              <dt>Source</dt>
              <dd>{sourceLabel}</dd>
              <dt>Format</dt>
              <dd>{info ? (FORMAT_LABELS[info.format] ?? info.format) : "—"}</dd>
              <dt>Slides / pages</dt>
              <dd>{info ? info.pageCount : "—"}</dd>
              <dt>Size</dt>
              <dd>
                {info ? `${(info.widthEmu / 914400).toFixed(2)}in × ${(info.heightEmu / 914400).toFixed(2)}in` : "—"}
              </dd>
              <dt>Aspect</dt>
              <dd>{info ? info.aspectRatio.toFixed(3) : "—"}</dd>
              <dt>Current slide</dt>
              <dd>{info ? `${slide + 1} / ${info.pageCount}` : "—"}</dd>
              <dt>Diagnostics</dt>
              <dd>{info ? info.diagnostics.length : "—"}</dd>
              <dt>Presenting</dt>
              <dd>{presenting ? "yes" : "no"}</dd>
              <dt>Error</dt>
              <dd>{error ? `${error.code}: ${error.message}` : "none"}</dd>
            </dl>
          </section>

          <section>
            <h3>Props in use</h3>
            <Code code={snippet} compact />
          </section>
        </aside>
      </div>

      <div className="pg__foot">
        <span className="pg__file">{fileName}</span>
        <span className="pg__note">
          {buildError ? (
            <span className="pg__error">{buildError}</span>
          ) : current ? (
            current.description
          ) : (
            "Your file is read in this tab only. An .html file is shown as source text and never executed."
          )}
        </span>
        <span className="pg__local">
          <i aria-hidden="true" /> Rendered locally
        </span>
        <a
          className="pg__report"
          href={reportHref}
          target="_blank"
          rel="noreferrer"
          title="Opens a GitHub bug report with the format and diagnostics filled in. Your file is not included."
        >
          <svg viewBox="0 0 20 20" aria-hidden="true">
            <path d="M4 17V3.5M4 4h9l-1.5 3L13 10H4" />
          </svg>
          Report an issue
        </a>
      </div>
    </div>
  );
}
