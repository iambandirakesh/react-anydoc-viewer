# Architecture

How `react-anydoc-viewer` turns a `.pptx`, `.docx` or `.pdf` file into pixels, and why
it is built this way. This is the design record for the decisions taken in phase 1
of the project (research and architecture selection); the code follows it.

The package was called `react-pptx-viewer` when presentations were all it opened.
The name is now format-neutral, and so is the public API — `DocumentViewer`,
`DocumentSource`, `DocumentInfo`, `DocumentViewerError`, `radv-`-prefixed CSS. Every
previous name is still exported as a deprecated alias of the same object (see
`src/index.ts`), because a rename should cost a find-and-replace, not a rewrite.

## Decisions up front

**Own parser instead of an existing renderer.** The open-source options were
evaluated before any code was written:

| Option | Verdict |
| --- | --- |
| `pptxjs` / `PPTXjs` | jQuery-era code, no build, no types, no React, unmaintained; not usable as a library dependency. |
| `pptx2html` | HTML conversion, not a viewer; no npm package or licence suitable for redistribution. |
| `docx-preview` | Renders WordprocessingML, not PresentationML. Its *approach* (client-side OOXML → HTML) informed the Word front-end added later, but not its code: this viewer paginates and draws into the same shape model as a slide. |
| `pptx-preview` | Vue-oriented, focused on conversion rather than a typed viewer API. |
| `PptxGenJS` | A **generator**, not a parser. Reading is out of scope. |
| Aspose / GroupDocs / Syncfusion | Commercial SDKs — they violate the "no paid API, no SDK licence" requirement. |
| LibreOffice / OnlyOffice | Server-side conversion; explicitly excluded. |
| pptx → PDF/image pipelines | Require a service. Excluded. |

No existing package offers a maintained, permissively licensed *parser* for
PresentationML, so that parser is implemented here, on top of two MIT/ISC
dependencies: `fflate` (ZIP) and `fast-xml-parser` (XML). The WordprocessingML
front-end is built the same way and shares everything downstream of the parse.
Everything else is standard browser APIs, SVG and CSS — no WebAssembly and no
rendering service (the only canvas in the project is a PDF page).

**PDF is the deliberate exception: reuse pdf.js.** A PDF is a program for a
rendering engine (fonts, images, blend modes, filters, colour spaces), not a shape
list, and a partial re-implementation would be both wrong and unbounded in scope.
So PDFs are handled by `pdfjs-dist` — Mozilla's engine, Apache-2.0, the code behind
Firefox's viewer — rasterised page by page onto a canvas. This is the one place the
project does not write the renderer itself, and it is what keeps "we render what the
file says" honest for this format. `pdfjs-dist` is a lazy, dynamically imported
dependency, so it is never downloaded by a `.pptx` or `.docx` user.

pdf.js ships a modern build and a `legacy` build that bundles polyfills, and an API
cannot be paired with the other build's worker. Which one runs is therefore decided
by feature detection (`detectPdfEngine()` checks for `Float16Array`,
`Uint8Array.prototype.toHex`, `Promise.try`, … rather than sniffing a user agent),
and the answer is frozen on first use — the legacy build's own polyfills would
otherwise make the next measurement report a modern environment. When a configured
`pdfWorkerSrc` turns out to be unusable (the wrong build, a URL that 404s, a blocked
cross-origin script), the file is still opened: the adapter retries in-process and
reports `pdf.engine.worker`.

**SVG for geometry, HTML for text.** PowerPoint shapes are vector geometry with
raw outlines, gradients, rotations and flips. SVG expresses that exactly and
scales losslessly. (A PDF page is the exception, because it is not geometry the
viewer interpreted: it is one canvas produced by pdf.js, positioned and scaled like
any other page box.) Text, however, must wrap, ellipsize, be selectable and reach the
accessibility tree, which is what HTML does well. Every shape therefore renders as
a small absolutely positioned box holding an SVG layer plus an HTML text layer,
and one CSS transform handles rotation/flipping for the whole box.

**The slide is never re-laid out.** The slide is rendered once at its natural
pixel size (EMU → px at 96 dpi) and scaled with `transform: scale()`, exactly like
zooming in PowerPoint. Text metrics therefore never change while zooming, and the
layout cost is paid once per slide.

## Pipeline

```
File | Blob | ArrayBuffer | Uint8Array | string | URL
        │
        ▼
  loadPresentation   fetch + size cap, abort handling, error normalisation
        │
        ▼
  format             magic bytes → pdf | zip | ole; part names → pptx | docx | other
        │
        ├─ pdf ─────────▶ pdf       pdf.js document + a page renderer
        │
        ▼ (zip)
  archive            ZIP index only (no inflate), safety limits, path validation
        │
        ▼
  relationships      .rels graph, r:id → target resolution, external vs internal
        │
        ├─ pptx ────────▶ presentation   theme → master → layout → slides, media
        │
        └─ docx ────────▶ docx   styles → numbering → blocks → measure → pages
        │
        ▼
  PresentationModel  plain data, no XML, no DOM
        │
        ▼
  React components   DocumentViewer → SlideCanvas → SlideView → RenderShape/RenderText/…
                     … or the format's own page renderer
                        (PdfPage, ImagePage, TextPage, GridPage)
```

The boundary that matters is the `PresentationModel`: the parsers produce plain
typed objects and the renderer consumes them, so the two halves are testable in
isolation (the whole parser suite runs without a DOM) and the renderer never needs
to know about XML namespaces, ZIP entries or relationship ids. That boundary is what
made the second and third formats cheap: a page of a Word document is a slide whose
shapes are paragraphs, tables and pictures, so navigation, zoom, thumbnails,
presentation mode and the error boundary came for free.

Two seams keep the exotic formats out of the rest of the code:

- **`probePackage(names)`** decides *which* parser runs, from part names only — no
  extension, MIME type or file name is consulted anywhere in the pipeline.
- **`PageRenderer`** (see `src/renderer/pageRenderer.ts`) lets a format draw a page
  itself instead of supplying shapes. `SlideCanvas` and `ThumbnailSidebar` call it
  when it is present, so exactly one component knows that a PDF page is a canvas,
  that an image is an `<img>`, that a text file is monospace lines and that a table
  is a grid, and the zoom scale is fed into the rasterisation so pages stay sharp.
  The context also carries the viewer's labels, so a renderer that wanted an
  accessible name of its own would not hard-code English.

Formats that are not OPC packages enter through the same entry point: an image is
identified by its magic bytes, a text file by the fact that it decodes as text
(with the *reading* — JSON, HTML, Markdown, delimited table, plain text — decided
from the content, because text has no signature). Both are dispatched before the
ZIP path. An HTML file is the same bytes as any other text file and takes the same
path: it is paginated as source, with a diagnostic naming what it is. No renderer
in the library ever executes file content, so there is no HTML frame to reason
about.

### Parser (`src/parser`)

| Module | Responsibility |
| --- | --- |
| `document.ts` | The single entry point every source funnels through: sniff the container, dispatch to the right parser, return `{ format, model, media, renderPage?, dispose }`. |
| `format.ts` | Magic-byte container detection (`%PDF-`, OLE/CFB, ZIP) and package-kind detection from part names, plus the wording of the "this is an `.xlsx`, not a document" errors. |
| `archive.ts` | Reads the ZIP central directory with `fflate`, enforces every limit, rejects unsafe entry names, sniffs image formats from magic bytes. Inflates lazily and only for the parts the renderer needs. |
| `relationships.ts` | Parses `.rels` parts and resolves ids relative to the source part, including `../` targets and `TargetMode="External"`. |
| `xml.ts` | Wraps `fast-xml-parser` in `preserveOrder` mode so element order (z-order, run order, `mc:AlternateContent`) survives; drops the XML declaration; rejects DOCTYPE/ENTITY. |
| `theme.ts`, `color.ts` | Theme colour scheme plus font scheme, and the `clrMap` indirection (`tx1` → `dk1` → `sysClr windowText`). Colour modifiers (`lumMod`, `tint`, `shade`, `alpha`, `satMod`, …) are applied numerically. |
| `drawing.ts` | The DrawingML vocabulary shared by shapes, theme and charts: transforms, fills, lines, shadows, geometry, adjustment values. |
| `geometry.ts` | 108 named presets plus `custGeom`, emitted as SVG path commands with a `fillRule`. |
| `text.ts` | Body properties (anchor, insets, autofit, wrap), paragraph properties, level styles, runs, bullets and numbering. |
| `shapes.ts` | The shape-tree walker: `p:sp`, `p:pic`, `p:graphicFrame`, `p:grpSp`, `p:cxnSp`, `mc:AlternateContent`, with group coordinate-space mapping and placeholder inheritance. |
| `table.ts`, `chart.ts` | Table grids (merges, per-edge borders) and chart families read from cached plot data. |
| `media.ts` | Lazy object URLs (`data:` URLs where object URLs are unavailable), one per media part, revoked on dispose. Shared byte-to-URL helper used by the image adapter. |
| `ole.ts` | Identifies a compound file from its stream names (`Workbook`, `WordDocument`, `PowerPoint Document`, `EncryptedPackage`), so legacy binary files are reported by name instead of as "encrypted". |
| `image.ts` | Header-only geometry for PNG/JPEG/GIF/BMP/WebP/SVG, and the one-page model for a standalone image. |
| `plaintext/` | The text front-end: `decode.ts` (BOM, UTF-8, UTF-16, Windows-1252 fallback), `classify.ts` (content-based classification and delimiter sniffing), `csv.ts` (quote-aware state machine), `paginate.ts` (wrapping and pagination) and `metrics.ts` (the page geometry and text metrics the page renderers import as well). |
| `../formats.ts` | The format table: labels, extensions and MIME types, from which `SUPPORTED_EXTENSIONS`, `ACCEPT_ATTRIBUTE` and every "supported formats are …" message are generated. |
| `presentation.ts` | Assembles everything: slide order from `sldIdLst`, layout/master resolution, inherited shape layers, background resolution, diagnostics. |
| `docx/` | The whole Word front-end: `styles.ts` (style chain, doc defaults, numbering), `text.ts` (paragraphs → run/picture/page-break blocks), `table.ts`, `drawing.ts` (inline and anchored pictures), `layout.ts` (pagination) and `index.ts` (orchestration). |
| `pdf.ts` | The pdf.js adapter: loads the engine lazily, reads page count, first-page geometry and metadata into a model, and returns a page renderer. Normalises pdf.js's own errors into the public error model. |

Parsing is synchronous after the download for the two OOXML formats and produces
one immutable model per source; the PDF path is asynchronous because pdf.js is
(and, unlike the OOXML parsers, it is allowed to be — the viewer already awaits the
load before committing it). Unsupported constructs become `Diagnostic` entries and,
where the renderer can, a labelled node — never a silent omission.

### Renderer (`src/renderer`, `src/components`)

| Module | Responsibility |
| --- | --- |
| `RenderShape.tsx` | Positioned box per shape: SVG geometry layer + HTML text layer, rotation/flips, groups, connectors, pictures with crop, unsupported placeholders. |
| `RenderText.tsx` | Paragraphs as absolutely placed blocks inside a flex column (vertical anchor), inline runs as spans, bullets as positioned spans. |
| `RenderTable.tsx` | Tables as absolutely positioned cells (merges produce the right spans). |
| `RenderChart.tsx` | Charts redrawn from cached series data as accessible inline SVG (axes, gridlines, legend, data labels). |
| `style.ts` | Shared mapping of model values to CSS/SVG (fills, gradients, dash arrays, shadows, ids). |
| `pageRenderer.ts` | The seam a non-shape format implements: given a page, its natural size, the current scale and a `warn` callback, return a node. Nothing else in the viewer knows PDFs exist. |
| `PdfPage.tsx` | The PDF implementation of that seam: sizes a canvas from the scale × device pixel ratio (capped for memory), renders one page with pdf.js, cancels on unmount, and turns a failure into a labelled notice plus a diagnostic. |
| `measureParagraph.ts` | Measures laid-out text with a hidden mirror DOM (`kind: "layout"`), and falls back to a metrics-based estimate (`kind: "estimate"`) where no DOM is available, for example in SSR or a test environment. |
| `DocumentViewer.tsx` | Public component: loading state machine, slide/zoom/fit state, keyboard handling, imperative handle, diagnostics, thumbnails, fullscreen. Format-agnostic: it forwards `renderPage` and maps the parsed model into `DocumentInfo`. |
| `SlideCanvas.tsx` | Scroll container, measurement, memoised slide view, per-slide error boundary — and the one line that delegates to a format's own page renderer. |
| `ViewerToolbar.tsx` | The chrome itself: a compact floating pill centred at the bottom of the stage, revealed by `:hover`/`:focus-within` in the stylesheet so it needs no visibility state of its own. Presentation only; every action and every piece of state belongs to `DocumentViewer`. |
| `ZoomMenu.tsx` | The zoom percentage and its dropdown of zoom stops and fit modes. The percentage and caret are one split button that opens on click or from the keyboard (never on hover), and it only calls back — it never owns the scale. |

Diagnostics reported during rendering (an undecodable image, for example) are
buffered and flushed in an effect after the commit, because a state update during
another component's render is not allowed by React.

## Security model

A `.pptx` is untrusted input, and the viewer is expected to run on user-uploaded
files:

1. **Archive**: entry-count, per-entry, total-size and compression-ratio limits are
   checked against the central directory *before* anything is inflated; entry names
   containing `..`, absolute paths, backslashes or control characters are ignored,
   and the parts that get inflated are restricted to a useful-parts allowlist.
2. **XML**: DOCTYPE and ENTITY declarations are refused outright, so entity
   expansion ("billion laughs") and external entities are not reachable. Text is
   decoded and unescaped manually for the five predefined entities only.
3. **Links**: `javascript:`, `data:`, `vbscript:` and unknowns are never emitted as
   `href`; they degrade to plain text with a `hyperlink.blocked` diagnostic.
4. **Images**: SVG parts are served through `<img>`/blob URLs, where scripts do not
   run. Media never becomes part of the host document.
5. **No HTML/JS execution from the file**: nothing in the pipeline evaluates
   document content. Text is rendered as text nodes, and an HTML file is read as
   text like any other: no iframe, no `srcdoc`, no `innerHTML`, so it has no origin
   to script from and no request is made on its behalf.
6. **Resource discipline**: object URLs are revoked on unmount and when the source
   changes; in-flight loads are aborted, and stale results are ignored.
7. **Third-party engine**: pdf.js runs with `isEvalSupported: false`, so a PDF
   cannot reach `eval`, and the page rasters it produces are capped (3× CSS density,
   8 megapixels) so a hostile page size cannot exhaust canvas memory. A PDF is
   handed to pdf.js as a copy of the bytes, because pdf.js transfers the buffer to
   its worker.
8. **One format registry, one dispatch**: the format is decided from the bytes, so
   there is no path where a file name influences which parser runs.

## Performance model

- The archive is indexed once; each reusable part (theme, master, layout, chart,
  style, numbering) is parsed at most once and memoised by part name.
- Only relevant parts are inflated — embedded workbooks, notes and diagrams are
  never decompressed.
- `pdfjs-dist` is behind a dynamic `import()`, so it is downloaded only when a PDF
  is opened, and only when the consumer's bundler code-splits. It never appears in
  the library bundle (`external` in the build).
- A `.docx` is paginated once, at load, by measuring each paragraph in a hidden
  mirror DOM; the measurement buffers are reused and removed afterwards.
- Media bytes are turned into URLs on first use, so a deck with 200 images only
  creates URLs for the slide being looked at (thumbnails are lazy-mounted by an
  `IntersectionObserver`).
- The visible slide is memoised; changing slides re-renders one slide.
- The renderer never measures the DOM to place content: all geometry comes from
  the file.

## Testing strategy

- **Fixtures are generated, not checked in**: `tests/fixtures/pptx.ts` writes real
  OOXML packages (masters, layouts, themes, slides, media, charts) with `fflate`,
  `tests/fixtures/docx.ts` writes WordprocessingML packages, and
  `tests/fixtures/pdf.ts` assembles a real PDF byte by byte (objects, xref, trailer,
  metadata). No proprietary file is redistributed, and every edge case (missing
  relationship, missing part, empty slide, unsupported media, unsafe hyperlink,
  renamed format, encrypted file) can be produced exactly.
- **Parser tests** run without a DOM.
- **Renderer and component tests** run under jsdom and exercise the component
  through the DOM the user sees: clicking toolbar buttons, pressing keys, reading
  the accessibility tree. PDF tests inject a fake pdf.js document through the
  `loadPdfJs` seam, which is why the adapter is typed structurally.
- **Browser tests** (`npm run test:browser`) run in real Chrome against the demo:
  they lay out a `.docx` into pages and rasterise a PDF page with pdf.js in a real
  Web Worker, which no jsdom test can prove.
- **Security tests** assert the archive limits, ZIP-bomb rejection, path traversal
  filtering, XML entity refusal and hyperlink sanitisation.
- **Package tests** (`npm run test:package`) build, `npm pack`, then install the
  tarball into a React 19 ESM app (which renders a real deck and asserts the slide
  text reaches the DOM) and a React 17 CJS app (SSR), plus a consumer type-check
  against the published declarations.

## Known architectural limits

- The OOXML parsers are synchronous on the main thread; only PDF parsing can be
  moved off it, and only because pdf.js brings its own worker. A worker-based
  implementation for `.pptx`/`.docx` would be the obvious next step for very large
  files, at the cost of a bundling story for the worker.
- Word pagination is the viewer's own, driven by measured text. It cannot match
  Word exactly while floating shapes, headers/footers and missing fonts are out of
  the model; the divergences are reported rather than hidden.
- PDFs are rasterised, so their text is not selectable and not in the
  accessibility tree — the deliberate cost of reusing a real engine instead of
  re-implementing one.
- The style cascade is deliberately shallow: placeholder geometry and defaults are
  inherited, but PowerPoint's full style matrix (`fmtScheme` references, table
  styles) is approximated.
- Thumbnails render the same component as the main view, so a 500-slide deck
  mounts slides on demand but still parses all of them once.
