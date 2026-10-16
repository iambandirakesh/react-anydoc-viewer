# react-anydoc-viewer

A free, open-source React component that renders **PowerPoint (`.pptx`)**, **Word
(`.docx`)**, **PDF**, **images**, **CSV/TSV tables** and **text files** directly in
the browser.

- **No watermark, no branding, no ads** — ever.
- **No paid API, no SDK licence, no external service.** Parsing and rendering happen entirely on the client.
- **One component for everything**: slide navigation, thumbnails, zoom and
  presentation mode do not care whether the file is a deck, a document, a PDF, a
  photograph, a spreadsheet extract or a log file.
- **No new dependencies for the new formats.** Images, CSV, text and JSON are
  handled with the ZIP/XML packages the library already had; pdf.js remains the
  only heavy optional download, and only for PDFs.
- **Small public API**: one component, one optional handle, a typed error model.
- **Presentation (slideshow) mode** built in — chrome-free playback with
  click/Space/arrow advance, for both mouse and keyboard audiences.
- **MIT licensed**, with MIT/ISC dependencies — plus Apache-2.0 for pdf.js, which
  is loaded lazily and only when a PDF is actually opened.

> **This is a viewer, not a PowerPoint/Word replacement.** It renders a
> deliberately documented subset of each format faithfully and reports everything
> it cannot render. See the [compatibility matrix](COMPATIBILITY.md) for exactly
> what is supported, partially supported and unsupported. Do not assume Microsoft
> Office parity — it has not been demonstrated and is not claimed.

---

## Table of contents

- [Installation](#installation)
- [Renamed from react-pptx-viewer](#renamed-from-react-pptx-viewer)
- [Basic usage](#basic-usage)
- [Supported formats](#supported-formats)
- [Loading documents](#loading-documents)
- [URLs and CORS](#urls-and-cors)
- [Customising the toolbar](#customising-the-toolbar)
- [Presentation (slideshow) mode](#presentation-slideshow-mode)
- [Imperative navigation](#imperative-navigation)
- [Error handling](#error-handling)
- [Diagnostics and blank-slide troubleshooting](#diagnostics-and-blank-slide-troubleshooting)
- [Styling and theming](#styling-and-theming)
- [Props reference](#props-reference)
- [Supported PowerPoint elements](#supported-powerpoint-elements)
- [Word documents](#word-documents)
- [PDF files](#pdf-files)
- [Performance](#performance)
- [Security](#security)
- [Browser support](#browser-support)
- [Known limitations](#known-limitations)
- [Development](#development)
- [Publishing](#publishing)
- [Licence and attribution](#licence-and-attribution)

---

## Installation

```bash
npm install react-anydoc-viewer
```

`react` and `react-dom` are **peer dependencies** (React 17 or newer) and are never bundled.

Import the stylesheet once, wherever your app is bootstrapped:

```ts
import "react-anydoc-viewer/styles.css";
```

### Renamed from `react-pptx-viewer`

The package was called `react-pptx-viewer` while it only opened presentations. The
name now reflects the formats it reads. **Existing code keeps working**: the
old exports are aliases, not copies, so they are the same component, class and
types — but they carry a `@deprecated` tag, and the CSS prefix changed from `rpv-`
to `radv-`.

| Old | New |
| --- | --- |
| `react-pptx-viewer` (npm, import path) | `react-anydoc-viewer` |
| `<PPTViewer />` | `<DocumentViewer />` (alias: `PPTViewer`) |
| `PPTViewerHandle`, `PPTViewerProps` | `DocumentViewerHandle`, `DocumentViewerProps` |
| `PPTViewerError`, `isPPTViewerError`, `PPTViewerErrorCode` | `DocumentViewerError`, `isDocumentViewerError`, `DocumentViewerErrorCode` |
| `PptxSource`, `PptxDiagnostic` | `DocumentSource`, `DocumentDiagnostic` |
| `PresentationInfo` | `DocumentInfo` |
| `rpv-*` classes, `--rpv-*` properties | `radv-*`, `--radv-*` |

Nothing else changed: the props, the imperative handle, the error codes and the
diagnostics are identical under both names, so a migration is a find-and-replace
whenever it suits you.

## Basic usage

```tsx
import { DocumentViewer } from "react-anydoc-viewer";
import "react-anydoc-viewer/styles.css";

export function App() {
  return (
    <DocumentViewer
      file="/example.pptx"
      showControls
      showThumbnails
      fit="width"
      onLoad={(info) => console.log(info.slideCount, "slides")}
      onError={(error) => console.error(error.code, error.message)}
      onSlideChange={(index) => console.log("slide", index)}
    />
  );
}
```

The viewer fills its parent element, so give the container a height:

```tsx
<div style={{ height: "80vh" }}>
  <DocumentViewer file="/example.pptx" />
</div>
```

## Supported formats

| Format | How it is rendered | Notes |
| --- | --- | --- |
| `.pptx` | Parsed here into a shape model, then drawn with SVG + HTML | Slides, themes, images, tables and charts |
| `.docx` | Parsed here, paginated in the browser, then drawn with the same shape renderer | A page is a slide as far as the viewer is concerned |
| `.pdf` | Rasterised page by page by [pdf.js](https://mozilla.github.io/pdf.js/) onto a canvas | The engine every browser PDF viewer uses |
| `.png` `.jpg` `.jpeg` `.gif` `.webp` `.avif` `.bmp` `.svg` | One page, sized from the file's own header | SVG included: an image is always an `<img>`, so it can never script your page |
| `.csv` `.tsv` | Parsed here, then drawn as a table with a repeated header row | Delimiter detected from the content; a wide table gets a landscape page |
| `.txt` `.log` `.json` `.html` `.htm` `.md` | Paginated as monospace text on US-Letter pages | JSON is pretty-printed; Markdown is currently shown as source too, and an HTML file is always shown as **source** — never parsed, never framed, never executed |

**The format is decided from the bytes, never from the file name.** `%PDF-`
identifies a PDF, the ZIP/OPC header identifies a package, and the package's own
part names (`ppt/presentation.xml`, `word/document.xml`) decide whether it is a
deck or a document. Images are identified by their magic bytes. Text files have no
signature at all, so the bytes decide *whether* a file is text and the content
decides *how* to read it (JSON, HTML, a consistent-delimiter table, otherwise plain
text) — pass `contentHint` when your application already knows.

A `.docx` renamed to `.pptx`, a file with no extension, or an `ArrayBuffer`
straight from `fetch` all load correctly. A spreadsheet (`.xlsx`) is refused with a
specific error instead of a guess, and a legacy binary `.doc`, `.ppt` or `.xls` is
named as such rather than being described as encrypted.

### HTML files are shown as source

An uploaded `.html` file is active content, so the viewer reads it and does nothing
else with it: the file is paginated as monospace text like any other text file, its
markup is never parsed, and it is never put in a frame — so there is no origin for
it to run a script from, no host document for it to reach, and no network request
made on its behalf. There is nothing to opt into: a `.html` file is displayed
through exactly the same path as a `.txt`, and a diagnostic says so.

If you want an uploaded page actually rendered, that belongs in your own
application, from a separate cookie-less origin — a stronger boundary than any
iframe attribute this library could set. [SECURITY.md](SECURITY.md) records exactly
what this covers and what it does not.

## Loading documents

`source` (or its alias `file`) accepts the natural browser inputs:

```tsx
// File from an <input type="file"> or drag-and-drop
<DocumentViewer source={file} />

// Bytes you already have
<DocumentViewer source={arrayBuffer} />
<DocumentViewer source={uint8Array} />

// A Blob from a fetch/Response
<DocumentViewer source={response.blob()} />

// A URL (string or URL object)
<DocumentViewer source="/decks/quarterly.pptx" />
<DocumentViewer source={new URL("/decks/quarterly.pptx", location.href)} />
```

A complete file-picker example:

```tsx
import { useState } from "react";
import { ACCEPT_ATTRIBUTE, DocumentViewer } from "react-anydoc-viewer";
import "react-anydoc-viewer/styles.css";

export function Uploader() {
  const [file, setFile] = useState<File | null>(null);
  return (
    <>
      <input
        type="file"
        accept={ACCEPT_ATTRIBUTE}
        onChange={(event) => setFile(event.target.files?.[0] ?? null)}
      />
      <div style={{ height: "70vh" }}>
        <DocumentViewer source={file} showControls showThumbnails />
      </div>
    </>
  );
}
```

Changing `source` to a different file cancels the previous load: an in-flight
download is aborted, object URLs from the previous document are revoked, and a late
result from the old file can never be committed to the screen.

## URLs and CORS

When you pass a URL the viewer calls `fetch`:

- The response must be a **CORS-enabled** same-origin URL, or served with
  `Access-Control-Allow-Origin`. Cross-origin requests that are not opted in fail
  with `NETWORK_ERROR`.
- Redirects, credentials and caching follow normal `fetch` semantics;
  `fetchOptions` is passed straight through, so cookies can be opted into with
  `fetchOptions={{ credentials: "include" }}`.
- Only `http:`, `https:`, `blob:` and `data:` URLs are fetched. Other schemes are
  rejected with `UNSUPPORTED_INPUT`.
- Failures surface as `NETWORK_ERROR` (including CORS failures, which browsers do
  not distinguish from network errors) or `LIMIT_EXCEEDED` when the response is
  larger than `maxDownloadBytes` (128 MiB by default).

> Private, cookie-authenticated documents served from another origin generally need
> `download`-then-`File` instead of direct URL loading: fetch the bytes yourself
> and pass the resulting `Blob` as `source`.

## Customising the toolbar

The toolbar is on by default and can be turned off entirely:

```tsx
<DocumentViewer source={file} showControls={false} />   // no chrome at all
<DocumentViewer source={file} showToolbar={false} />    // alias of showControls
<DocumentViewer source={file} showThumbnails />         // opens the sidebar and its toggle
<DocumentViewer source={file} showThumbnails allowThumbnailToggle={false} />
<DocumentViewer source={file} showPresentationControl={false} />  // hide the Present button
```

With `showControls={false}` there is no Present button; drive the slideshow from
your own UI through the handle instead — see
[Presentation mode](#presentation-slideshow-mode).

The chrome is a **small pill floating at the bottom of the stage**, centred on the
slide, not a header bar: the document keeps the full width of the viewer and the
slide is scaled to fit in the space left above the pill.

It also **stays out of the way while reading**. The pill is hidden until the viewer
is hovered, and in that hidden state it is inert (`pointer-events: none`), so it can
never swallow a click meant for the document. It is not removed from the DOM: focus
brings it back, so the keyboard path is the same either way. On a touch device
(where there is no hover) it is always visible.

Everything in the pill is short by design — previous/next, the counter, zoom out,
the zoom-and-fit dropdown, zoom in, Present, thumbnails and fullscreen. Jumping
straight to the first or last page is a keyboard (`Home`/`End`) and handle action
rather than two more buttons in every reader's view, and Present is an icon whose
full text ("Present from this slide") is its accessible name and tooltip.

Zoom and fit share a single dropdown on that pill — click the percentage (it does
not open on hover, so the pointer crossing the pill never drops a menu over the
slide):

```
      ‹  2 / 8  ›  │  −  ┌ 67% │ ▾ ┐  +  │  ▶  ▤  ⛶
                         └─────┴───┘
                              │
                      ┌───────┴────────┐
                      │ 10%  25%  50%  │   zoom stops
                      │ 75%  100% 125% │
                      │ …              │
                      ├────────────────┤
                      │ Fit width  ✓   │   fit modes
                      │ Fit height     │
                      │ Fit page       │
                      │ Actual size    │
                      └────────────────┘
```

The percentage and the caret are one button — a divider between them, one
accessible name, one tab stop — and it opens on click or on `Enter`, `Space`,
`ArrowUp` and `ArrowDown`. It is a `menu` of `menuitemradio` entries, so the mode in
effect is announced to assistive technology: arrows and `Home`/`End` move through
the entries, and `Esc` closes it and returns focus to the percentage. Picking a percentage is a zoom (`onZoomChange`); picking a fit mode
is a fit (`onFitChange`).

Labels are fully localisable, including the slide counter, which is a function:

```tsx
<DocumentViewer
  source={file}
  labels={{
    previous: "Vorherige Folie",
    next: "Nächste Folie",
    slideOf: (current, total) => `Folie ${current} von ${total}`,
  }}
/>
```

For a completely custom UI, hide the chrome, drive navigation through a ref and
render your own buttons. Loading and error states can be replaced with
`renderLoader` and `renderError`.

### Keyboard shortcuts

Enabled by default (`keyboardNavigation`). The viewer is a focusable region; keys
are handled while it or one of its controls has focus.

| Key | Action |
| --- | --- |
| `→` `↓` `PageDown` `Space` `n` | Next slide |
| `←` `↑` `PageUp` `p` | Previous slide |
| `Home` / `End` | First / last slide |
| `+` / `-` | Zoom in / out |
| `0` | Reset to the `fit` mode |
| `f` | Toggle fullscreen |

## Presentation (slideshow) mode

Presentation mode hides every control, fills the viewport with the slide on a black
background and starts from whatever slide is currently shown:

```tsx
<DocumentViewer source={file} defaultPresentation />        {/* start presenting */}
<DocumentViewer source={file} presentation={yes} onPresentationChange={setYes} />
<DocumentViewer source={file} showPresentationControl={false} />   {/* never show the button */}
```

The toolbar gets a **Present from this slide** button (an icon, labelled in full for
assistive technology and on hover) unless `showPresentationControl={false}`. While
presenting:

- **click**, `Space`, `Enter`, `→`, `↓`, `PageDown` or `n` advance;
- `←`, `↑`, `PageUp`, `Backspace` or `p` go back; `Home`/`End` jump to the ends;
- `Esc` exits — always, even if the fullscreen overlay is also on — and the small
  hint at the bottom carries an **Exit presentation** button for pointer-only
  users (it fades after a few seconds and returns when the pointer moves);
- the toolbar, thumbnail sidebar and diagnostics panel are removed from the DOM, so
  nothing behind the slide is focusable;
- leaving restores the previous chrome, sidebar, zoom and fit mode untouched.

Press `f` inside a slideshow to use the browser's fullscreen as well (which also
hides the address bar); `enterPresentation()` and `enterFullscreen()` compose.

## Imperative navigation

```tsx
import { useRef } from "react";
import { DocumentViewer, type DocumentViewerHandle } from "react-anydoc-viewer";

function Controlled() {
  const viewer = useRef<DocumentViewerHandle>(null);
  return (
    <>
      <button onClick={() => viewer.current?.goToSlide(4)}>Go to slide 5</button>
      <button onClick={() => viewer.current?.setZoom(2)}>200%</button>
      <div style={{ height: "60vh" }}>
        <DocumentViewer ref={viewer} source="/deck.pptx" />
      </div>
    </>
  );
}
```

The handle exposes `goToSlide`, `next`, `previous`, `first`, `last`,
`getCurrentSlide`, `getSlideCount`, `setZoom`, `zoomIn`, `zoomOut`, `resetZoom`,
`getZoom`, `enterFullscreen`, `exitFullscreen`, `toggleFullscreen`,
`enterPresentation`, `exitPresentation`, `togglePresentation`, `isPresenting` and
`reload`.

For controlled usage, pass `slide` (0-based) and/or `zoom` and observe
`onSlideChange` / `onZoomChange`.

## Error handling

Every failure is normalised into a `DocumentViewerError` with a stable `code`, so you
never have to match on message text:

```tsx
import { isDocumentViewerError } from "react-anydoc-viewer";

<DocumentViewer
  source={file}
  onError={(error) => {
    if (error.code === "ENCRYPTED") showHint("Remove the password in PowerPoint first");
  }}
  renderError={(error, retry) => (
    <div>
      <p>{error.message}</p>
      <code>{error.code}</code>
      <button onClick={retry}>Try again</button>
    </div>
  )}
/>
```

| Code | Meaning |
| --- | --- |
| `UNSUPPORTED_INPUT` | `source` was not a supported input type or scheme |
| `NETWORK_ERROR` | Download failed (network, CORS or HTTP status) |
| `NOT_A_PRESENTATION` | The bytes are not a document of any supported format |
| `UNSUPPORTED_FORMAT` | A real Office package, but not a format the viewer renders (`.xlsx`, OpenDocument, an unfamiliar OPC package) |
| `ENCRYPTED` | Password-protected, or the legacy binary `.ppt`/`.doc` format |
| `LIMIT_EXCEEDED` | A configured archive/download limit was exceeded |
| `CORRUPT_PACKAGE` | The ZIP directory or an entry could not be read |
| `NO_SLIDES` | No slide could be resolved from the presentation |
| `MALFORMED_PRESENTATION` | The presentation part is missing or unreadable |
| `MALFORMED_DOCUMENT` | A damaged PDF, or a `.docx` without a readable `word/document.xml` |
| `RENDER_ERROR` | A slide threw while rendering (contained per slide) |
| `ABORTED` | A newer load superseded this one |
| `UNKNOWN` | Anything else; see `error.detail` |

Failures that affect only part of a deck are **not** errors: they are reported as
diagnostics (see below) and the rest of the presentation still renders.

## Diagnostics and blank-slide troubleshooting

A blank-looking slide should always be explicable. The viewer distinguishes three
cases, and never shows a silent white rectangle for a failed render:

1. **A genuinely empty page** renders a "This page is empty" notice (custom
   text via `emptySlideMessage`, or `labels.emptySlide`).
2. **An element that cannot be rendered** (SmartArt, an OLE object, an EMF/WMF
   image, an unimplemented preset shape) renders a small labelled placeholder in
   place of the element. A PDF page that pdf.js cannot rasterise behaves the same
   way and reports `pdf.page`.
3. **A page that throws while rendering** is contained by an error boundary and
   shows "This page failed to render" plus the message, without taking down the
   viewer. Replace it with `renderSlideError`.

To see everything the parser and renderer skipped, turn on `debug` (or subscribe
with `onDiagnostics`):

```tsx
<DocumentViewer source={file} debug />
```

The diagnostics panel lists every message with its severity, a stable code (for
example `picture.unsupported-format`, `geometry.preset-unsupported`,
`hyperlink.blocked`, `table.style-approximated`) and the slide it belongs to.

Other common causes of an apparently blank viewer:

| Symptom | Cause |
| --- | --- |
| "Waiting for the container to have a usable size…" | The parent element has zero height/width |
| Nothing at all, `data-status="idle"` | No `source` was passed |
| Empty-looking slide | The slide has no shapes and no background — see case 1 above |
| Text missing but shapes present | A font that is not installed, or text inherited from a layout placeholder that the deck left empty |

## Styling and theming

Import the packaged stylesheet, then override what you need. All classes are
`radv-` prefixed and the palette comes from CSS custom properties:

```css
.my-viewer {
  --radv-chrome: #101418;
  --radv-chrome-border: #232a31;
  --radv-text: #f2f5f7;
  --radv-accent: #6ea8fe;
  --radv-stage: #05070a;
}
```

```tsx
<DocumentViewer source={file} className="my-viewer" dark />
```

`dark` (or `theme="dark"`) applies the built-in dark chrome. Pass `style` for
inline tweaks, or skip `showControls` and build your own dashboard around the
component.

## Props reference

| Prop | Type | Default | Description |
| --- | --- | --- | --- |
| `source` / `file` | `File \| Blob \| ArrayBuffer \| Uint8Array \| string \| URL` | — | The document to display: `.pptx`, `.docx` or `.pdf` |
| `slide` | `number` | — | Controlled current slide (0-based) |
| `defaultSlide` | `number` | `0` | Uncontrolled initial slide |
| `showControls` | `boolean` | `true` | Show the toolbar |
| `showToolbar` | `boolean` | — | Alias of `showControls` |
| `showThumbnails` | `boolean` | `false` | Show the thumbnail sidebar |
| `allowThumbnailToggle` | `boolean` | `showThumbnails` | Show the sidebar toggle button |
| `presentation` | `boolean` | — | Controlled presentation (slideshow) mode |
| `defaultPresentation` | `boolean` | `false` | Start in presentation mode |
| `showPresentationControl` | `boolean` | `true` | Show the Present button in the toolbar |
| `fit` | `"width" \| "height" \| "page" \| "actual" \| "none"` | `"page"` | Initial scaling mode |
| `zoom` | `number` | — | Controlled scale (`1` is 100%) |
| `defaultZoom` | `number` | — | Uncontrolled initial scale |
| `keyboardNavigation` | `boolean` | `true` | Enable keyboard shortcuts |
| `theme` | `"light" \| "dark"` | `"light"` | Viewer chrome theme |
| `dark` | `boolean` | `false` | Shorthand for `theme="dark"` |
| `className` / `style` | `string` / `CSSProperties` | — | Applied to the root element |
| `emptySlideMessage` | `string` | `"This page is empty"` | Notice for a genuinely empty page |
| `labels` | `Partial<ViewerLabels>` | — | Localised strings |
| `debug` | `boolean` | `false` | Show the diagnostics panel |
| `limits` | `Partial<ArchiveLimits>` | see below | Archive safety limits |
| `fetchOptions` | `RequestInit` | — | Passed to `fetch` for URL sources |
| `maxDownloadBytes` | `number` | `134217728` | Maximum accepted download size |
| `contentHint` | `"text" \| "csv" \| "json" \| "html" \| "markdown"` | — | How to read an ambiguous text file; skips the content guess |
| `pdfWorkerSrc` | `string` | — | URL of pdf.js's worker module; match it with `detectPdfEngine()` |
| `onLoad` | `(info: DocumentInfo) => void` | — | Called after a successful parse |
| `onError` | `(error: DocumentViewerError) => void` | — | Called on any load/parse failure |
| `onSlideChange` | `(index: number) => void` | — | Called when the visible slide changes |
| `onZoomChange` | `(scale: number) => void` | — | Called when the zoom scale changes |
| `onFitChange` | `(fit: FitMode) => void` | — | Called when the fit mode changes |
| `onPresentationChange` | `(active: boolean) => void` | — | Called when presentation mode starts or ends |
| `onDiagnostics` | `(diagnostics: DocumentDiagnostic[]) => void` | — | Receives parse diagnostics |
| `renderLoader` | `() => ReactNode` | — | Replace the loading indicator |
| `renderError` | `(error, retry) => ReactNode` | — | Replace the error state |
| `renderSlideError` | `(error, slideIndex) => ReactNode` | — | Replace the per-slide failure panel |

`ArchiveLimits` defaults (all configurable through `limits`):

| Limit | Default | Purpose |
| --- | --- | --- |
| `maxCompressedBytes` | 128 MiB | Maximum input size |
| `maxEntries` | 8000 | Maximum archive entry count |
| `maxEntryBytes` | 128 MiB | Maximum uncompressed size per entry |
| `maxTotalBytes` | 512 MiB | Maximum total uncompressed size |
| `maxCompressionRatio` | 200 | Decompression-bomb heuristic |
| `maxXmlBytes` | 48 MiB | Maximum size of a single XML part |

## Supported PowerPoint elements

The full breakdown is in [COMPATIBILITY.md](COMPATIBILITY.md). In short:

**Implemented** — slide size and aspect ratio, backgrounds (solid, gradient,
picture), theme colour schemes with the master colour map, theme font schemes,
text with runs, paragraph alignment, indentation, line spacing, bullets and
numbered lists, bold/italic/underline/strikethrough, colour and highlight,
superscript/subscript, shapes from 108 named preset geometries plus custom
geometry paths, solid/gradient/picture fills, outlines with dash patterns, outer shadows,
images with cropping and non-uniform scaling, rotation and flipping, groups with
nested coordinate spaces, tables with merges/fills/per-edge borders, bar, line,
area, pie, doughnut and scatter charts from cached plot data, hyperlinks, and
slide layouts/master background graphics.

**Partial** — inherited placeholder text styles are resolved through master and
layout styles rather than a complete style cascade; table cell styling
(`tableStyles.xml`) is approximated with a neutral grid; 3-D chart variants render
as their 2-D equivalent; `mc:AlternateContent` uses the fallback branch.

**Unsupported** — SmartArt, embedded OLE objects, animations and transitions,
video/audio playback, EMF/WMF/TIFF images, 3-D shapes and effects, and the legacy
binary `.ppt` format.

## Word documents

A `.docx` is not a page of shapes, so the viewer parses WordprocessingML itself
(`word/document.xml`, `styles.xml`, `numbering.xml`, the relationship and media
parts) exactly as it parses PresentationML, then **paginates the document in the
browser** and hands each page to the same renderer a slide uses. Navigation,
thumbnails, zoom, fit modes and presentation mode therefore work unchanged, and
`slideCount` is the page count.

| Supported | Notes |
| --- | --- |
| Page size and margins (`w:pgSz`, `w:pgMar`), portrait and landscape | Custom page sizes are honoured; the last section's setup applies to the whole document |
| Paragraphs: alignment, indents, spacing before/after, line spacing, borders, shading | Twips, half-points and EMU are converted to CSS pixels |
| Runs: bold, italic, underline, strike, colour, highlight, size, font, super/subscript, caps | |
| Styles: named paragraph and character styles, `basedOn` inheritance, doc defaults | Built-in heading/list styles are recognised |
| Bulleted and numbered lists from `numbering.xml`, including nesting | `lvlText`, formats (`decimal`, `roman`, `alpha`, `bullet`) and restarts |
| Tables: grid columns, `gridSpan` (merges), row heights, cell shading, borders | Nested tables are reported, not drawn |
| Inline pictures, including extent-driven scaling | Floating/anchored shapes and text wrapping are reported |
| Hyperlinks, with the same scheme allowlist as slides | |
| Explicit page breaks and keep-with-next | |
| `docProps/core.xml` title and author | Exposed on `onLoad` |

Reported rather than rendered: headers and footers, footnotes and endnotes,
comments, equations (`m:oMath`), text boxes and other floating shapes, revision
marks (insertions are kept, deletions skipped), embedded (`altChunk`) content, and
section breaks, which are treated as page breaks with a diagnostic. Page breaks are
computed from **measured** text — the browser gives the real line heights — with an
estimated fallback (reported as `docx.layout.estimated`) where no DOM is available.

## PDF files

PDFs are handed to [pdf.js](https://mozilla.github.io/pdf.js/) — the engine behind
Firefox's own viewer. The library does not re-implement a PDF rasteriser, and that
is deliberate: a PDF page is a program for an engine, not a shape list. Each page is
rasterised onto a canvas at the viewer's current scale (times the device pixel
ratio), so it stays sharp when zoomed, and only the visible page and visible
thumbnails are ever drawn. Password-protected PDFs report `ENCRYPTED`; a damaged one
reports `MALFORMED_DOCUMENT`.

`pdfjs-dist` is a normal dependency (Apache-2.0). It is **not** part of the library
bundle and is not downloaded at all unless a PDF is opened, so `.pptx` and `.docx`
users pay nothing for it. The build that runs is chosen by feature detection — the
modern one where it can run, pdf.js's `legacy` build (polyfills included) everywhere
else — and the choice is reported as `pdf.engine.legacy` in the diagnostics.

### The PDF worker (recommended)

pdf.js normally parses in a Web Worker. A published component cannot invent a
worker URL inside someone else's bundler, so by default the viewer tells pdf.js to
parse on the main thread — that works everywhere the module resolves, and reports a
`pdf.worker.main-thread` notice in the diagnostics. For production, publish the
worker and point the viewer at it:

```tsx
// Vite / webpack 5 / Parcel: copy the worker into your output and give it a URL
const pdfWorkerSrc = new URL(
  "pdfjs-dist/build/pdf.worker.min.mjs",
  import.meta.url,
).href;

<DocumentViewer source={file} pdfWorkerSrc={pdfWorkerSrc} />;
```

```tsx
// Next.js: import the worker as an asset and pass the resulting URL
<DocumentViewer source={file} pdfWorkerSrc={pdfWorkerUrl} />;
```

**Publish the worker that matches the engine.** pdf.js ships two builds and refuses
to pair an API with the other build's worker, while the modern build assumes
features (`Uint8Array.prototype.toHex`, `Float16Array`, …) that older Chromium
shells — Electron applications, for example — do not have. The viewer feature-detects
the build it needs; `detectPdfEngine()` tells you which one, so the asset you publish
agrees with it:

```tsx
import { detectPdfEngine } from "react-anydoc-viewer";

const pdfWorkerSrc =
  detectPdfEngine() === "legacy"
    ? new URL("pdfjs-dist/legacy/build/pdf.worker.min.mjs", import.meta.url).href
    : new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url).href;
```

If the worker you passed turns out to be unusable — the other build, a URL that
404s, a blocked cross-origin script — the file still opens: the parse is retried
in-process and `pdf.engine.worker` is reported instead of failing.

Without a worker everything still works, but a page's first paint can stutter on
large files because parsing shares the main thread. The viewer itself renders a page
only when it becomes visible, so a 500-page PDF does not parse 500 pages up front.

## Performance

- The archive and every reusable part (theme, master, layout, chart, table style)
  are parsed **at most once** per document.
- Only the parts the renderer needs are inflated; embedded workbooks, notes and
  diagrams are never decompressed.
- Media bytes are turned into object URLs **on demand**, when a page actually
  uses them, and are revoked on unmount or when a new source loads.
- The visible page is rendered at its natural size and scaled with a CSS
  transform, so zooming never re-flows text and scrolling does not re-render.
- Thumbnails are mounted only when they scroll near the viewport, so a 200-slide
  deck does not render 200 slides up front.
- Rendering is memoised per page, so changing pages does not re-render the rest.

The `.pptx` and `.docx` parsers run synchronously on the main thread after the
download (the load is `await`, so React paints the loading state first). Most decks
parse in a few milliseconds; a very large deck (hundreds of slides with large
embedded media) will block briefly during that step. A `.docx` additionally measures
each paragraph, which costs a layout pass per page — still milliseconds for ordinary
documents. PDFs are parsed by pdf.js, in a worker when `pdfWorkerSrc` is set.

Bundle size (measured on the built artefacts): **86.3 kB gzipped** for the minified
ESM bundle — the ZIP/XML dependencies, all three format front-ends and the viewer —
plus **2.3 kB gzipped** of CSS. `pdfjs-dist` is excluded from those numbers because
it is loaded lazily. Nothing is loaded from a CDN and no fonts or images are
requested by the library itself.

## Security

Presentation files are treated as untrusted input:

- **Archive limits** — entry count, per-entry size, total size and a
  decompression-ratio heuristic are enforced before anything is inflated.
- **Path traversal** — entry names containing `..`, absolute paths, backslashes
  or control characters are ignored.
- **XML** — DOCTYPE/ENTITY declarations are refused, entity expansion is disabled,
  and only the five predefined entities plus numeric character references are
  decoded (with a size cap). Element depth is bounded by input-size limits.
- **Scripts** — the viewer never evaluates embedded JavaScript or HTML. Slide text
  is rendered as text. SVG parts are only ever used as an `<img>` source, where
  browsers do not run scripts.
- **HTML files** — shown as source text, paginated like any other text file, and
  never put in a frame: no origin, no scripts, no network loads. There is no option
  to render them. See [SECURITY.md](SECURITY.md).
- **Links** — hyperlinks are validated with the URL parser and restricted to
  `http:`, `https:`, `mailto:` and `tel:`. `javascript:`, `data:`, `vbscript:`,
  `file:` and credential-bearing URLs are dropped (and reported as
  `hyperlink.blocked`).
- **External media** — linked (non-embedded) images and other external
  relationships are never fetched; the viewer makes no network requests on behalf
  of a document's content.
- **No injection surface** — nothing from the file is interpolated into HTML,
  `innerHTML`, `dangerouslySetInnerHTML`, or a style attribute without escaping.
- **No scripting or eval from content** — document text is React text nodes. An
  external hyperlink is only ever an `href` after the scheme allowlist, and pdf.js
  is configured with `isEvalSupported: false`, so a PDF cannot reach `eval`.
- **Third-party engine** — PDFs are parsed by `pdfjs-dist` (Mozilla, Apache-2.0),
  which is left to do what it does in Firefox; no part of a file is passed to any
  service. Keep it up to date like any other dependency.
- **Bounded rendering** — page rasters are capped (3x CSS density, 8 megapixels
  per page) so a hostile page size cannot exhaust canvas memory.

See [SECURITY.md](SECURITY.md) for the threat model and how to report an issue.

## Browser support

Modern evergreen browsers: Chrome/Edge 90+, Firefox 90+, Safari 15+.

Requirements: ES2020, `fetch`, `URL.createObjectURL`, `ResizeObserver` (with a
`window.resize` fallback), CSS flexbox, inline SVG. No polyfills are shipped.
PDFs additionally need a 2D canvas and a Web Worker when `pdfWorkerSrc` is set;
without a canvas the viewer shows a labelled notice instead of a blank page.
`IntersectionObserver` and the Fullscreen API are optional — thumbnails fall back
to rendering eagerly, and fullscreen falls back to the viewer's own
fixed-position overlay (`data-fullscreen="overlay"`) so the control always works,
including in Electron-style webviews and iframes without an `allowfullscreen`
policy. `Esc` leaves the overlay.

Server rendering works for the component's own markup (the initial state renders
without touching the DOM); pages need a real DOM, and the media store falls back
to `data:` URLs when object URLs are unavailable. `pdfjs-dist` is imported lazily
inside the effect that loads a PDF, so it is never touched during SSR.

## Known limitations

- **Not a complete PowerPoint/Word replacement.** Complex files will differ from
  Office. Unsupported constructs are labelled, never silently dropped.
- Fonts are whatever the browser has: decks and documents that rely on Calibri,
  Cambria or a custom corporate font will fall back. Ship or load the font in your
  app if fidelity matters.
- SmartArt, OLE objects, animations, transitions and media playback are not
  rendered.
- EMF/WMF/TIFF images cannot be decoded by browsers; they render as labelled
  placeholders.
- Speaker notes are not parsed or displayed.
- `mc:AlternateContent` renders the fallback branch, which is the conservative
  choice but can differ from modern PowerPoint output.
- **Word**: pagination is the viewer's own, so page breaks can differ from Word by
  a line when a floating shape, a header/footer, or an unusual font is involved.
  Headers, footers, footnotes, comments and equations are reported, not rendered;
  a multi-section document uses its last section's page setup throughout.
- **PDF**: rendering is pdf.js's, not a pixel-exact copy of Acrobat, and forms,
  annotations, JavaScript and portfolios are shown as their static appearance.
  The accessibility tree gets one image per page, not the text layer.
- Text auto-fit (`normAutofit`) applies the recorded font scale but does not
  re-run PowerPoint's fitting algorithm, so a very long line can still overflow.
- There is no print, no export and no editing.
- Parsing is synchronous on the main thread except for PDFs (see
  [Performance](#performance)).
- Native fullscreen depends on the embedding environment. Where the browser
  blocks or ignores `requestFullscreen()` (some desktop webviews, sandboxed
  iframes), the viewer expands into its own overlay instead; the visual result is
  the same but it does not use the browser's fullscreen state.
- Text layout is the browser's, not PowerPoint's. Line breaking can differ by a
  pixel or two, and a paragraph that PowerPoint would shrink can overflow instead.

## Development

```bash
npm install
npm run dev          # demo app on http://localhost:5273
npm run test         # unit, parser, renderer and component tests (Vitest)
npm run typecheck    # TypeScript, no emit
npm run build        # ESM + CJS + UMD bundles, styles.css, .d.ts
npm run test:package # build, npm pack, then verify from separate React 19 and 17 apps
npm run test:browser # Playwright integration tests in real Chrome (starts the demo)
```

The browser suite drives the demo application end to end: navigation, zoom and fit
modes, thumbnails, keyboard shortcuts, fullscreen (native or overlay), diagnostics,
the empty-slide notice, chart geometry, resizing, a broken upload and a
no-console-errors pass over every slide — plus a Word document laid out into pages
and each PDF page actually rasterised onto a canvas by pdf.js with a real worker. It
launches the Chrome already installed on the machine (`channel: "chrome"`); on a
machine without Chrome, run `npx playwright install chromium` and drop that option.

The demo app builds every sample in the browser (see
[examples/demo/src/sampleDeck.ts](examples/demo/src/sampleDeck.ts),
[examples/demo/src/sampleDocument.ts](examples/demo/src/sampleDocument.ts)), so no
binary files are committed — including the PDF, which is assembled byte by byte. Open
<http://localhost:5273> and press **Load sample deck**, **Load sample .docx** or
**Load sample .pdf**; the toggles try the thumbnail sidebar, dark chrome, `debug`
diagnostics and each fit mode. Dropping a `.pptx`, `.docx` or `.pdf` onto the stage
loads your own file.

Note: developing the demo requires Node 20.19+ or 22.12+ (a Vite requirement).
The published package itself only requires Node 18+ for tooling that loads it.

The test suite generates its own `.pptx` fixtures at runtime
(`tests/fixtures/pptx.ts` writes real OOXML packages), so no third-party or
proprietary presentations are redistributed with this project.

Architecture notes live in [ARCHITECTURE.md](ARCHITECTURE.md); the feature matrix
in [COMPATIBILITY.md](COMPATIBILITY.md); release history in
[CHANGELOG.md](CHANGELOG.md).

## Publishing

The package is prepared for npm but is **not** published automatically, and nothing
is published by any workflow in this repository. Release it from your own machine,
once per version:

```bash
# 0. One-time: authenticate as the account that owns the name.
npm login              # opens the browser; 2FA prompts are handled there
npm whoami             # must print your npm username

# 1. Confirm the name is still yours to take (404 = free).
npm view react-anydoc-viewer version

# 2. Verify from a clean install
npm ci
npm run typecheck
npm test
npm run test:package   # packs the tarball into React 19 and React 17 consumer apps

# 3. Rehearse the upload. This is the real thing minus the PUT: it runs
#    prepublishOnly (typecheck + tests + build), then prints every file and the
#    tarball size. Nothing reaches the registry.
npm publish --dry-run

# 4. Publish. 0.1.0 has never been on the registry, so it goes out as-is.
npm publish --access public

# 5. Verify the live artefact (the registry can lag a few seconds)
npm view react-anydoc-viewer version dist.tarball

# 6. For every release after this one: bump semver, add the CHANGELOG entry,
#    then publish again. `0.1.0` itself is never bumped — a published version
#    number cannot be reused, so that step belongs to the *next* release.
npm version patch      # or minor / major
npm publish --access public
```

What the upload contains, measured on `0.1.0`: **118 files, ~617 kB packed,
2.1 MB unpacked** — `dist/` (ESM, CJS, UMD, `styles.css`, `.d.ts`), `package.json`,
`LICENSE` and the Markdown docs. Sources, tests, examples and the demo are excluded
by the `files` field in `package.json`.

Notes:

- **The name is not reserved until the first publish.** `react-anydoc-viewer` was
  unclaimed on the public registry when this package was prepared (`npm view` returns
  a 404). Re-check in step 1; if it has been taken, either pick another name in
  `package.json` or publish under a scope (`@your-scope/react-anydoc-viewer`, and
  mirror the same value in your import paths).
- **`prepublishOnly` runs the type-checker, the test suite, the build and a
  publish gate** (`scripts/prepublish-guard.mjs`, offline), so a broken package
  cannot be published by accident — including under `--dry-run`, which makes step 3
  a complete verification pass. The gate fails the upload if ownership metadata
  still points at a placeholder, if `files` would ship sources or tests, or if
  `main`/`exports` promise a file the build never wrote.
- **Signatures and provenance.** npm signs ordinary publishes and shows the
  "Provenance" badge only for builds run from a public repository in CI (for
  example GitHub Actions with `npm publish --provenance` and an `id-token: write`
  permission). A local publish is signed by the registry, but carries no provenance
  attestation.
- **`npm version` requires a git repository to commit and tag.** Without one it only
  rewrites the version in `package.json`, so commit and tag the release yourself
  (`git tag -a v0.1.0 -m "0.1.0"`) before publishing.
- **A published version is permanent.** You may unpublish a whole package within 72
  hours of its first publish, but an individual version number can never be reused
  once it has been published — a bad release is fixed with a new patch version, not
  by editing `0.1.0`.
- **Keep the npm account secured** with 2FA ("auth-and-writes") and, for automation,
  a granular access token scoped to this package rather than a classic token.
- Test the artefact before and after publishing with `npm run test:package`, which
  installs the packed tarball into separate React 19 and React 17 projects.

## Licence and attribution

MIT — see [LICENSE](LICENSE). You may use it in commercial products, modify it and
redistribute it, with no attribution requirement beyond keeping the licence text.

Bundled runtime dependencies:

| Package | Licence | Purpose |
| --- | --- | --- |
| [fflate](https://github.com/101arrowz/fflate) | MIT | Reading the ZIP/OPC package |
| [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser) | MIT | XML parsing with order preservation |
| [pdfjs-dist](https://github.com/mozilla/pdf.js) | Apache-2.0 | Rasterising PDF pages |

All three are permissive and compatible with redistributing this package under MIT
(pdf.js's Apache-2.0 is not copyleft). No proprietary SDK, paid API or external
conversion service is used or required.

Contributions are welcome — please run `npm run typecheck && npm test` before
opening a pull request. Test presentations must be generated fixtures or
otherwise legally distributable; do not add proprietary decks.
