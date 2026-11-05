/**
 * Demo site.
 *
 * A landing page around the live playground: what the library is, a viewer to try
 * every format in, what it supports, and how to install it. Every sample is
 * generated in the browser, so the site makes no document requests at all.
 */

import { useEffect, useState, type ReactNode } from "react";
import { version } from "react-anydoc-viewer";
import { Playground } from "./Playground";
import { Code, useCopy } from "./Code";
import { GITHUB_URL } from "./report";

const NPM_URL = "https://www.npmjs.com/package/react-anydoc-viewer";
const INSTALL = "npm install react-anydoc-viewer";

const USAGE = `import { DocumentViewer } from "react-anydoc-viewer";
import "react-anydoc-viewer/styles.css";

export function Preview({ file }: { file: File }) {
  // The viewer fills its parent, so give the parent a height.
  return (
    <div style={{ height: "80vh" }}>
      <DocumentViewer file={file} />
    </div>
  );
}`;

const CONTROLLED = `const viewer = useRef<DocumentViewerHandle>(null);

<DocumentViewer
  ref={viewer}
  file="/reports/q3.pptx"
  fit="width"
  onLoad={(info) => console.log(info.format, info.pageCount)}
  onSlideChange={setPage}
/>

viewer.current?.togglePresentation();`;

type Theme = "light" | "dark";

function readTheme(): Theme {
  if (typeof document === "undefined") return "light";
  return document.documentElement.dataset.theme === "dark" ? "dark" : "light";
}

function Icon({ d }: { d: string }) {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d={d} />
    </svg>
  );
}

const FEATURES: { title: string; body: string; icon: string }[] = [
  {
    title: "Nothing leaves the browser",
    body: "Files are parsed and drawn on the device. No server, no API key, no conversion service and no upload.",
    icon: "M12 3 4 6v6c0 4.5 3.4 8.3 8 9 4.6-.7 8-4.5 8-9V6zM9 12l2 2 4-4",
  },
  {
    title: "One component, every format",
    body: "Slides, Word, Excel, PDF, images, CSV and text share one toolbar, zoom, thumbnails, keyboard map and API.",
    icon: "M4 5h7v7H4zM13 5h7v7h-7zM4 14h7v5H4zM13 14h7v5h-7z",
  },
  {
    title: "Free, with no watermark",
    body: "MIT licensed. No licence key, no trial banner, no usage limit — ship it in commercial apps as it is.",
    icon: "M12 2v20M17 6H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6",
  },
  {
    title: "Safe with untrusted files",
    body: "The format is read from the bytes, archives have size limits, SVG is an <img> and HTML is only ever shown as source.",
    icon: "M5 11V8a7 7 0 0 1 14 0v3M4 11h16v10H4zM12 15v2",
  },
  {
    title: "Presentation mode",
    body: "A full-screen slideshow for decks, documents and sheets, driven by click, keyboard or your own controls.",
    icon: "M3 4h18v12H3zM8 20h8M12 16v4M10 8l4 2-4 2z",
  },
  {
    title: "Light on your bundle",
    body: "pdf.js stays out of the bundle and is fetched only when a PDF is opened. ESM, CJS and UMD for React 17 to 19.",
    icon: "M13 2 4 14h7l-1 8 9-12h-7z",
  },
];

const FORMATS: { name: string; badge: string; tagline: string; exts: string[]; tone: string; points: string[] }[] = [
  {
    name: "PowerPoint",
    badge: "PPTX",
    tagline: "Slides and decks",
    exts: [".pptx"],
    tone: "orange",
    points: ["Themes, text and shapes", "Images, groups and tables", "Charts from cached data", "Thumbnails and slideshow"],
  },
  {
    name: "Word",
    badge: "DOCX",
    tagline: "Paginated documents",
    exts: [".docx"],
    tone: "blue",
    points: ["Paginated in the browser", "Headings, lists and tables", "Pictures and page breaks", "Same viewer as slides"],
  },
  {
    name: "Excel",
    badge: "XLSX",
    tagline: "Styled workbooks",
    exts: [".xlsx", ".xlsm"],
    tone: "green",
    points: ["Fonts, fills, borders, merges", "Number formats and dates", "Charts, images, data bars", "Search, copy, CSV export"],
  },
  {
    name: "PDF",
    badge: "PDF",
    tagline: "Rendered by pdf.js",
    exts: [".pdf"],
    tone: "red",
    points: ["Mozilla’s PDF engine", "Sharp at every zoom level", "Loaded only when needed", "Optional worker"],
  },
  {
    name: "Images",
    badge: "IMG",
    tagline: "Raster, vector and scans",
    exts: [".png", ".jpg", ".gif", ".webp", ".avif", ".bmp", ".svg", ".tiff"],
    tone: "violet",
    points: ["Sized from the file header", "Multi-page TIFF on canvas", "SVG can never run script", "Zoom and fit modes"],
  },
  {
    name: "Tables & text",
    badge: "CSV",
    tagline: "Delimited data and plain text",
    exts: [".csv", ".tsv", ".txt", ".json", ".md", ".log", ".html"],
    tone: "slate",
    points: ["Delimiter detected from content", "JSON pretty-printed", "HTML shown as source", "Encoding detection"],
  },
];

function HeroArt() {
  return (
    <div className="art" aria-hidden="true">
      <div className="art__card art__card--sheet">
        <div className="art__label">
          <b className="tone-green">XLSX</b> revenue.xlsx
        </div>
        <div className="art__grid">
          {Array.from({ length: 24 }, (_, index) => (
            <i key={index} className={index < 4 ? "h" : index % 4 === 3 ? "n" : ""} />
          ))}
        </div>
      </div>
      <div className="art__card art__card--doc">
        <div className="art__label">
          <b className="tone-blue">DOCX</b> proposal.docx
        </div>
        <div className="art__lines">
          <i className="t" />
          <i />
          <i />
          <i className="s" />
          <i />
          <i className="s" />
        </div>
      </div>
      <div className="art__card art__card--slide">
        <div className="art__label">
          <b className="tone-orange">PPTX</b> launch.pptx
        </div>
        <div className="art__slide">
          <div className="art__title" />
          <div className="art__bars">
            <i style={{ height: "38%" }} />
            <i style={{ height: "62%" }} />
            <i style={{ height: "48%" }} />
            <i style={{ height: "84%" }} />
            <i style={{ height: "70%" }} />
          </div>
        </div>
        <div className="art__toolbar">
          <i />
          <span>3 / 12</span>
          <i />
        </div>
      </div>
    </div>
  );
}

function Section({ id, eyebrow, title, lead, children }: { id: string; eyebrow: string; title: string; lead?: ReactNode; children: ReactNode }) {
  return (
    <section id={id} className="section">
      <div className="section__head">
        <p className="eyebrow">{eyebrow}</p>
        <h2>{title}</h2>
        {lead ? <p className="lead">{lead}</p> : null}
      </div>
      {children}
    </section>
  );
}

export function App() {
  const [theme, setTheme] = useState<Theme>(readTheme);
  const [copied, copy] = useCopy();

  useEffect(() => {
    document.documentElement.dataset.theme = theme;
    try {
      localStorage.setItem("radv-demo-theme", theme);
    } catch {
      // Storage can be unavailable (private mode); the theme then lasts the visit.
    }
  }, [theme]);

  return (
    <div className="site">
      <header className="nav">
        <div className="nav__inner">
          <a className="brand" href="#top">
            <span className="brand__mark" aria-hidden="true" />
            react-anydoc-viewer
          </a>
          <nav className="nav__links" aria-label="Sections">
            <a href="#playground">Playground</a>
            <a href="#features">Features</a>
            <a href="#formats">Formats</a>
            <a href="#start">Quick start</a>
            <a className="nav__report" href={`${GITHUB_URL}/issues/new/choose`} target="_blank" rel="noreferrer">
              Report an issue
            </a>
          </nav>
          <div className="nav__actions">
            <button
              type="button"
              className="icon-btn"
              onClick={() => setTheme(theme === "dark" ? "light" : "dark")}
              aria-label={theme === "dark" ? "Switch to light theme" : "Switch to dark theme"}
            >
              <Icon
                d={
                  theme === "dark"
                    ? "M12 4V2M12 22v-2M4 12H2M22 12h-2M5.6 5.6 4.2 4.2M19.8 19.8l-1.4-1.4M5.6 18.4l-1.4 1.4M19.8 4.2l-1.4 1.4M12 7a5 5 0 1 0 0 10 5 5 0 0 0 0-10z"
                    : "M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"
                }
              />
            </button>
            <a className="btn btn--ghost btn--sm" href={NPM_URL} target="_blank" rel="noreferrer">
              npm
            </a>
            <a className="btn btn--ghost btn--sm" href={GITHUB_URL} target="_blank" rel="noreferrer">
              <svg viewBox="0 0 16 16" aria-hidden="true" className="gh">
                <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0 0 16 8c0-4.42-3.58-8-8-8z" />
              </svg>
              GitHub
            </a>
          </div>
        </div>
      </header>

      <main id="top">
        <section className="hero">
          <div className="hero__copy">
            <a className="pill" href={`${GITHUB_URL}/blob/main/CHANGELOG.md`} target="_blank" rel="noreferrer">
              <span className="pill__dot" /> v{version} · MIT · open source
            </a>
            <h1>
              Every document.
              <br />
              <span className="grad">One React component.</span>
            </h1>
            <p className="hero__lead">
              Render PowerPoint, Word, Excel, PDF, images, CSV and text files right inside your React app. Parsed and
              drawn in the browser — no server, no API key, no watermark.
            </p>
            <div className="hero__cta">
              <a className="btn btn--primary" href="#playground">
                Try it live
                <Icon d="M5 12h14M13 6l6 6-6 6" />
              </a>
              <button type="button" className="install" onClick={() => copy(INSTALL)} aria-label={`Copy: ${INSTALL}`}>
                <span className="install__prompt">$</span>
                <code>{INSTALL}</code>
                <span className="install__hint">{copied ? "Copied" : "Copy"}</span>
              </button>
            </div>
            <ul className="chips" aria-label="Supported formats">
              {["PPTX", "DOCX", "XLSX", "PDF", "TIFF", "PNG · JPG · SVG", "CSV", "TXT · JSON"].map((label) => (
                <li key={label}>{label}</li>
              ))}
            </ul>
          </div>
          <HeroArt />
        </section>

        <Section
          id="playground"
          eyebrow="Playground"
          title="Try it with a sample, or your own file."
          lead="Every sample below is generated in this tab. Drop in any supported file — it is read locally and never uploaded."
        >
          <Playground siteDark={theme === "dark"} />
        </Section>

        <Section id="features" eyebrow="Why this viewer" title="Built for real apps, not just demos.">
          <div className="features">
            {FEATURES.map((feature) => (
              <article key={feature.title} className="feature">
                <span className="feature__icon">
                  <Icon d={feature.icon} />
                </span>
                <h3>{feature.title}</h3>
                <p>{feature.body}</p>
              </article>
            ))}
          </div>
        </Section>

        <Section
          id="formats"
          eyebrow="Formats"
          title="Format support at a glance."
          lead={
            <>
              The format is decided from the file’s bytes, never its name. The full matrix lives in{" "}
              <a href={`${GITHUB_URL}/blob/main/COMPATIBILITY.md`} target="_blank" rel="noreferrer">
                COMPATIBILITY.md
              </a>
              .
            </>
          }
        >
          <div className="formats">
            {FORMATS.map((format) => (
              <article key={format.name} className={`format tone-${format.tone}`}>
                <header className="format__head">
                  <span className="format__badge" aria-hidden="true">
                    {format.badge}
                  </span>
                  <div>
                    <h3>{format.name}</h3>
                    <p className="format__tagline">{format.tagline}</p>
                  </div>
                </header>
                <ul className="format__exts" aria-label="File extensions">
                  {format.exts.map((ext) => (
                    <li key={ext}>{ext}</li>
                  ))}
                </ul>
                <ul className="format__points">
                  {format.points.map((point) => (
                    <li key={point}>{point}</li>
                  ))}
                </ul>
              </article>
            ))}
          </div>
        </Section>

        <Section id="start" eyebrow="Quick start" title="Up and running in a minute.">
          <div className="start">
            <div className="start__steps">
              <div className="step">
                <span className="step__n">1</span>
                <div>
                  <h3>Install</h3>
                  <p>One package with everything it needs. pdf.js is loaded only when a PDF is opened.</p>
                  <Code code={INSTALL} title="terminal" />
                </div>
              </div>
              <div className="step">
                <span className="step__n">2</span>
                <div>
                  <h3>Render a file</h3>
                  <p>
                    Pass a <code>File</code>, <code>Blob</code>, <code>ArrayBuffer</code> or URL. The viewer fills its
                    parent.
                  </p>
                  <Code code={USAGE} title="Preview.tsx" />
                </div>
              </div>
            </div>
            <div className="step step--side">
              <span className="step__n">3</span>
              <div>
                <h3>Drive it from your app</h3>
                <p>Controlled props, callbacks and an imperative handle for navigation, zoom and presenting.</p>
                <Code code={CONTROLLED} title="controlled.tsx" />
              </div>
            </div>
          </div>
        </Section>

        <section className="cta">
          <h2>Show documents, not download links.</h2>
          <p>Free and open source under the MIT licence.</p>
          <div className="hero__cta">
            <a className="btn btn--primary" href={GITHUB_URL} target="_blank" rel="noreferrer">
              Star on GitHub
            </a>
            <a className="btn btn--ghost" href={`${GITHUB_URL}#readme`} target="_blank" rel="noreferrer">
              Read the docs
            </a>
          </div>
        </section>
      </main>

      <footer className="footer">
        <span>
          <span className="brand__mark brand__mark--sm" aria-hidden="true" /> react-anydoc-viewer · MIT licence
        </span>
        <span className="footer__links">
          <a href={GITHUB_URL} target="_blank" rel="noreferrer">
            GitHub
          </a>
          <a href={NPM_URL} target="_blank" rel="noreferrer">
            npm
          </a>
          <a href={`${GITHUB_URL}/blob/main/CHANGELOG.md`} target="_blank" rel="noreferrer">
            Changelog
          </a>
          <a href={`${GITHUB_URL}/issues/new/choose`} target="_blank" rel="noreferrer">
            Report an issue
          </a>
        </span>
      </footer>
    </div>
  );
}
