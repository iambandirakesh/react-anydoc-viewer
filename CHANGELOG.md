# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project follows
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

While the major version is `0`, minor versions may contain breaking changes; they
will be called out under **Changed**.

## [Unreleased]

### Performance

- A long `.docx` no longer costs one forced layout per paragraph. Every paragraph
  in a document is measured at the same width, so the whole document is now laid
  out in one pass and the heights are read afterwards: a 6,000-paragraph document
  (360 pages) loads in 624 ms instead of 727 ms on the machine the benchmark runs
  on, with the same pagination. A measurer that cannot batch is still called per
  paragraph, with identical results.
- `npm run bench` (Playwright) measures time-to-first-page for a generated
  `.pptx`, `.docx` and `.pdf`, scales them with `BENCH_SCALE`, and prints the CPU
  hotspots with `BENCH_PROFILE=1`, so a slow document can be attributed rather
  than guessed at.

## [0.1.0] — unreleased (initial release)

First public release. A free, watermark-free React viewer for `.pptx`, `.docx` and
`.pdf` files that parses and renders everything locally in the browser.

### Renamed

The package was called `react-pptx-viewer`, from when it only opened presentations.
Since the same component now renders `.docx` and `.pdf` as well, the name is
format-neutral: **`react-anydoc-viewer`**, with `DocumentViewer` as the component
and `DocumentSource`/`DocumentInfo`/`DocumentDiagnostic`/`DocumentViewerError` as
the types.

Every previous name is still exported as a deprecated alias — the *same* component
and the *same* class, so `instanceof`, refs, props and error codes keep working:
`PPTViewer`, `PPTViewerProps`, `PPTViewerHandle`, `PPTViewerError`,
`isPPTViewerError`, `PPTViewerErrorCode`, `PPTViewerErrorType`, `PptxSource`,
`PptxDiagnostic`, `PresentationInfo`. The CSS prefix and custom properties moved
from `rpv-`/`--rpv-` to `radv-`/`--radv-`; nothing is published under the old name,
so no consumer is left on the wrong side of that change.

### Added

**Public API**

- `DocumentViewer` component with `source`/`file` (File, Blob, ArrayBuffer, Uint8Array,
  URL or string), `showControls`, `showThumbnails`, `allowThumbnailToggle`, `fit`,
  controlled/uncontrolled `slide` and `zoom`, `keyboardNavigation`, `theme`/`dark`,
  `debug`, `limits`, `fetchOptions`, `maxDownloadBytes` and `labels` (localisable,
  including the slide counter as a function).
- `onLoad`, `onError`, `onSlideChange`, `onZoomChange`, `onFitChange`,
  `onDiagnostics`, plus `renderLoader`, `renderError` and `renderSlideError`.
- `pdfWorkerSrc` for publishing pdf.js's worker; without it PDFs parse on the main
  thread and the viewer reports `pdf.worker.main-thread`.
- `DocumentViewerHandle` for programmatic navigation and zoom, including
  `enterFullscreen`, `exitFullscreen` and `reload`.
- Typed error model (`DocumentViewerError` with a stable `code`) and typed diagnostics.
- `DocumentInfo.format` reports which of the three formats was detected, and
  `SUPPORTED_FORMATS` is exported for UI that needs the list.

**Formats**

- **Format detection from bytes, never from a name or MIME type.** The `%PDF-`
  header, the OLE/CFB signature and the ZIP/OPC signature are sniffed directly, and a
  package's kind is decided by its own part names (`ppt/presentation.xml`,
  `word/document.xml`). A renamed file, an extension-less file and an `ArrayBuffer`
  from `fetch` all open correctly; an `.xlsx`, an OpenDocument file or an
  unfamiliar package is refused with `UNSUPPORTED_FORMAT` naming what it actually is.
- **Word documents (`.docx`).** A WordprocessingML front-end (`word/document.xml`,
  `styles.xml`, `numbering.xml`, relationships, media) that builds the *same*
  `PresentationModel` the slide renderer already consumes, after paginating the
  document in the browser from measured text. Page setup, paragraphs, runs, styles
  with `basedOn` inheritance, lists, tables with merges, inline and anchored
  pictures, hyperlinks, page breaks and keep-with-next are rendered; headers,
  footers, footnotes, comments, equations, text boxes, VML/OLE content and
  `altChunk` are reported. Because a document page is a slide, navigation,
  thumbnails, zoom, fit modes, presentation mode and the error boundary work
  unchanged, and `slideCount` is the page count.
- **PDF files.** Pages are rasterised by `pdfjs-dist` (Mozilla, Apache-2.0) through
  a `PageRenderer` seam, at the viewer's current scale × device pixel ratio and
  capped for canvas memory, so pages stay sharp while zooming and only visible pages
  and thumbnails are drawn. Page count, first-page geometry, mixed page sizes and
  document metadata are read into the model. Password-protected PDFs report
  `ENCRYPTED`, damaged ones `MALFORMED_DOCUMENT`. `pdfjs-dist` is a lazy dependency:
  it is not part of the library bundle and is never loaded for a `.pptx` or `.docx`.
  The engine runs with `isEvalSupported: false` and no annotation or JavaScript
  layer, so page content cannot execute anything.
- **The pdf.js build is chosen by feature detection** (`detectPdfEngine()`, exported
  for hosts that publish the worker), because the modern build needs primitives such
  as `Uint8Array.prototype.toHex` and `Float16Array` that older Chromium shells do
  not implement; the legacy build's polyfills are used instead, reported as
  `pdf.engine.legacy`. A configured `pdfWorkerSrc` that cannot be used — the wrong
  build, a 404, a blocked cross-origin script — no longer fails the load: the parse
  is retried in-process and reported as `pdf.engine.worker`.

**Parser**

- OPC/ZIP reading with lazy inflation, slide order resolved through
  `p:sldIdLst` relationships, and theme → master → layout → slide inheritance.
- Text: runs, paragraph properties, level styles, spacing, bullets and
  auto-numbering, theme fonts and colours with DrawingML colour modifiers.
- Shapes: 108 preset geometries plus custom geometry, solid/gradient/picture
  fills, outlines with dash patterns, rotation and flips, groups with nested
  coordinate spaces, connectors.
- Pictures with cropping and non-uniform scaling; SVG media.
- Tables with merges, fills and per-edge borders; bar, line, area, pie, doughnut
  and scatter charts from cached plot data.
- Hyperlinks, backgrounds (solid, gradient, picture) and inherited layout/master
  background graphics.

**Viewer**

- Chrome rendered as a compact floating pill centred at the bottom of the stage
  rather than a header bar, so the slide keeps the full width of the viewer:
  previous/next, the counter, zoom in/out, the zoom **and fit** dropdown, Present,
  the thumbnail toggle and fullscreen, plus keyboard shortcuts and an accessible,
  `radv-`-prefixed DOM with `aria-live` status. The pill is hidden — and inert —
  until the viewer is hovered or one of its controls has focus, and is always
  visible where there is no hover to reveal it (touch devices).
- One dropdown for zoom and fit, listing the conventional zoom stops (10 %–500 %)
  and fit width/height/page and actual size together. Its trigger is a split
  control — the percentage, a divider, a caret — and opens on click or from the
  keyboard, uses `menuitemradio` semantics so the mode in effect is announced, and
  never covers the slide: the fit calculation reserves the pill's height.
- `fit="height"`, to scale a page to the height of the container.
- Presentation (slideshow) mode: `presentation` / `defaultPresentation` /
  `onPresentationChange` / `showPresentationControl`, plus `enterPresentation`,
  `exitPresentation`, `togglePresentation` and `isPresenting` on the handle. The
  chrome is removed from the DOM, the slide fills a black stage, clicks and
  Space/arrows advance, and `Esc` (or the hint's button) exits and restores the
  previous chrome, sidebar, zoom and fit mode.
- Thumbnail sidebar, lazily mounted with `IntersectionObserver`.
- Light and dark chrome via CSS custom properties; the slide itself always renders
  light, like a document.

**Diagnostics and blank-slide prevention**

- Diagnostics panel (`debug`) and `onDiagnostics` with stable codes for every
  construct the parser or renderer skipped.
- Distinct handling for an empty slide, an unrenderable element (labelled
  placeholder) and a slide that throws (per-slide error boundary), so a blank slide
  is never unexplained.
- Zero-size container detection with an explanatory notice.

**Security**

- Archive limits (entries, sizes, compression ratio) enforced before inflation,
  path-traversal filtering, DOCTYPE/ENTITY refusal, hyperlink scheme allowlist,
  SVG served through `<img>` so it cannot script the page, object-URL lifecycle
  management and abort handling for superseded loads.

**Packaging and docs**

- ESM + CJS + UMD builds, TypeScript declarations, `styles.css` subpath export,
  React `>=17` peer dependency, MIT licence.
- `npm run test:package` builds, packs, and verifies the tarball from separate
  React 19 (ESM) and React 17 (CJS/SSR) applications, plus a consumer type-check.
- 151 unit/component tests (parser, renderer, components, security) built on
  generated `.pptx`, `.docx` and `.pdf` fixtures — no third-party documents are
  redistributed — plus 17 Playwright integration tests that drive the demo in real
  Chrome, including a Word document laid out into pages and a PDF page rasterised
  with pdf.js in a real Web Worker.
- An empty-slide notice is no longer rendered (as an empty `role="note"` node) in
  thumbnails, where the message is deliberately suppressed.
- Demo application (`npm run dev`) that generates its own eight-slide deck, a Word
  document and a PDF in the browser, since no binaries are committed.
- Documentation: `README.md`, `COMPATIBILITY.md`, `ARCHITECTURE.md`,
  `SECURITY.md`, this changelog.

### Changed during 0.1.0 development

- **The chrome moved from a header bar to a small floating bottom pill, and it is
  only shown on hover or focus.** Fit width, fit page and actual size were three
  buttons in the header; they are now entries in the zoom dropdown, which is also
  where a percentage can be picked directly. The first/last jump buttons left the
  chrome to keep it narrow — `Home`/`End` and `first()`/`last()` on the handle do
  the same thing — and Present is now an icon carrying its full text as the
  accessible name and tooltip. Nothing was removed from the API: `fit`,
  `onFitChange` and the `FitMode` values still work, and `ViewerLabels` gained
  `fitHeight` and `zoomMenu`.
- **The zoom dropdown no longer opens on hover.** It appeared when the pointer
  rested on the percentage, which meant the pointer merely crossing the pill while
  reading dropped a 15-item menu over the slide. The percentage and its caret now
  read as one boxed control and open only on click (or `Enter`/`Space`/arrows).
- **A localised presentation hint could spill out of its own pill.** The hint row
  is `white-space: nowrap`, and its text span had no `min-width: 0`, so a flex item
  could not shrink below its content: a longer translation (or a longer
  `labels.exitPresentation`) overflowed the pill instead of ellipsising. The text
  now shrinks and ellipsises, and the pill clips whatever is left.
- **Default label copy is format-neutral**, since the same chrome now serves slides,
  document pages and PDF pages: "Loading document…", "This document could not be
  displayed", "This page is empty", "This page failed to render", "Page thumbnails".
  All of them remain overridable through `labels` (and `emptySlideMessage`).

### Fixed during 0.1.0 development

- **Fullscreen did nothing where the Fullscreen API is blocked.** Environments that
  report `document.fullscreenEnabled === true` but never settle
  `requestFullscreen()` (desktop webviews, sandboxed iframes) left the button dead
  and the rejection was swallowed. The viewer now watches the request and falls
  back to a fixed-position overlay (`data-fullscreen="overlay"`) that the toolbar,
  the `f` shortcut and `Esc` all control.
- **Paragraph margins came from the user agent.** A `margin: 0` shorthand was
  dropped by React 19 when longhand margins were also present, so every paragraph
  inherited the user agent's `1em` margins and text drifted out of its box. Margins
  are now always written as longhands, with a stylesheet reset for block elements
  inside the slide.
- **Every pie chart was drawn as a doughnut.** `holeSize` defaulted to 50 for all
  pie-like charts instead of only doughnut and pie-of-pie charts.
- **Render-time diagnostics updated parent state during render**, which React
  reports as "cannot update a component while rendering a different component" and
  which concurrent rendering may drop. Reports are now buffered and flushed after
  the commit.
- **Effects that CSS cannot express were dropped silently.** Glow, soft edge,
  reflection, preset shadow, fill overlay and 3-D bevel now produce
  `effect.unsupported`, and an inner shadow reports that it is approximated.
- **An image that failed to decode showed a broken-image icon** with no
  explanation; it now becomes a labelled placeholder and reports `picture.decode`.
- Percentage line spacing in the test fixtures was written in the wrong unit
  (`spcPct` is thousandths of a percent).
