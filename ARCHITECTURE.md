# Architecture

How `react-anydoc-viewer` turns presentations, documents, workbooks, PDFs, images,
CSV/TSV tables and text files into pixels. This describes the implementation through
0.2.3, including the worksheet view introduced after the original architecture
selection. Historical proposals are kept separately in
[docs/FORMAT-EXPANSION.md](docs/FORMAT-EXPANSION.md).

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
dependencies: `fflate` (ZIP) and `fast-xml-parser` (XML). The WordprocessingML and
SpreadsheetML front-ends are built the same way and share everything downstream of
the parse. Everything else is standard browser APIs, SVG and CSS — no WebAssembly
and no rendering service. Canvases appear only where a format *is* already a raster:
a PDF page from pdf.js and a TIFF page from the lazy decoder.

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
  loadDocument       fetch + size cap, abort handling, error normalisation
        │
        ▼
  format             magic bytes → pdf | zip | ole | image | text; ZIP parts → pptx | docx | xlsx
        │
        ├─ raster ──────▶ image     header geometry + image or TIFF page content
        ├─ pdf ─────────▶ pdf       pdf.js document as page content
        ├─ text ────────▶ plaintext CSV → one worksheet; other text → pages
        │
        ▼ (zip)
  archive            ZIP index only (no inflate), safety limits, path validation
        │
        ▼
  relationships      .rels graph, r:id → target resolution, external vs internal
        │
        ├─ pptx ────────▶ pptx   theme → master → layout → slides, media
        ├─ docx ────────▶ docx   styles → numbering → blocks → measure → pages
        └─ xlsx ────────▶ xlsx   shared strings → styles → worksheets + drawings
        │
        ▼
  DocumentModel      plain data, no XML, no DOM
  + PageContent      for self-drawn pages: plain data, mapped to a renderer by
                     components/pageRenderers.tsx
        │
        ▼
  React components   DocumentViewer → SlideCanvas → SlideView → RenderShape/RenderText/…
                     … or the format's own page renderer
                        (PdfPage, TiffPage, ImagePage, TextPage)
                     … or worksheet tabs/tools → SheetPage (XLSX and CSV)
```

The boundary that matters is the `DocumentModel`: the parsers produce plain
typed objects and the renderer consumes them, so the two halves are testable in
isolation (the whole parser suite runs without a DOM). The dependency only points
one way: nothing under `src/parser` imports from `src/renderer`, `src/components`
or `src/hooks`. The one piece of parsing that needs the DOM — measuring DOCX
paragraph heights — is injected through the parse options by the viewer and the renderer never needs
to know about XML namespaces, ZIP entries or relationship ids. That boundary is what
made the second and third formats cheap: a page of a Word document is a slide whose
shapes are paragraphs, tables and pictures, so navigation, zoom, thumbnails,
presentation mode and the error boundary came for free.

Shared parsing/rendering seams keep most format-specific work contained:

- **`probePackage(names)`** decides *which* parser runs, from part names only — no
  extension, MIME type or file name is consulted anywhere in the pipeline.
- **`PageContent`** (see `src/parser/pageContent.ts`) lets a format describe pages
  it draws itself instead of supplying shapes — a pdf.js document, TIFF bytes, an
  image URL, text pages or worksheets. `components/pageRenderers.tsx` maps each
  kind to a **`PageRenderer`** (`src/renderer/pageRenderer.ts`), which
  `SlideCanvas` and `ThumbnailSidebar` call for paginated formats. Sheet content
  (XLSX/CSV) makes `DocumentViewer` use `WorkbookView` — bottom tabs, internal
  scrolling and `WorksheetTools` — instead. Raster renderers receive the display scale to size their canvases.
  The context also carries the viewer's labels, so a renderer that wanted an
  accessible name of its own would not hard-code English.

Formats that are not OPC packages enter through the same entry point: an image is
identified by its magic bytes, a text file by the fact that it decodes as text
(with the *reading* — JSON, HTML, Markdown, delimited table, plain text — decided
from the content, because text has no signature). Both are dispatched before the
ZIP path. An HTML file is the same bytes as any other text file and takes the same
path: it is paginated as source, with a diagnostic naming what it is. No renderer
in the library evaluates document scripts or formula text, so there is no HTML
frame to reason about.

### Parser (`src/parser`)

| Module | Responsibility |
| --- | --- |
| `document.ts` | The single entry point every source funnels through: sniff the container, dispatch to the right parser, return `{ format, model, media, content?, dispose }`. |
| `format.ts` | Magic-byte container detection (`%PDF-`, OLE/CFB, ZIP) and package-kind detection from part names, plus the wording of the "this is an `.xlsb`, not a document" errors. |
| `archive.ts` | Reads the ZIP central directory with `fflate`, enforces every limit, rejects unsafe entry names, sniffs image formats from magic bytes. Inflates lazily and only for the parts the renderer needs. |
| `relationships.ts` | Parses `.rels` parts and resolves ids relative to the source part, including `../` targets and `TargetMode="External"`. |
| `xml.ts` | Wraps `fast-xml-parser` in `preserveOrder` mode so element order (z-order, run order, `mc:AlternateContent`) survives; drops the XML declaration; rejects DOCTYPE/ENTITY. |
| `drawingml/` | The DrawingML vocabulary shared by all three OOXML formats (PPTX slides, DOCX pictures, XLSX drawings): `theme.ts`/`color.ts` (theme colour and font schemes, the `clrMap` indirection, numeric colour modifiers), `drawing.ts` (transforms, fills, lines, shadows), `geometry.ts` (108 named presets plus `custGeom` as SVG paths), `text.ts` (body/paragraph/run properties, bullets), `shapes.ts` (the shape-tree walker with group mapping and placeholder inheritance), `table.ts` and `chart.ts`. |
| `media.ts` | Lazy object URLs (`data:` URLs where object URLs are unavailable), one per media part, revoked on dispose. Shared byte-to-URL helper used by the image adapter. |
| `ole.ts` | Identifies a compound file from its stream names (`Workbook`, `WordDocument`, `PowerPoint Document`, `EncryptedPackage`), so legacy binary files are reported by name instead of as "encrypted". |
| `image.ts` | Header-only geometry for PNG/JPEG/GIF/BMP/WebP/SVG, the TIFF path (IFD pages, drawn by `TiffPage`), and the one-page model for a standalone image. |
| `tiff.ts` | Bounded TIFF directory walking: page count and natural size from the IFD chain, without reading pixels. |
| `xlsx/` | `index.ts` assembles worksheets/limits; `styles.ts` handles numeric/date formats; `visualStyles.ts` maps fonts/fills/borders/alignment/theme colors; `drawings.ts` resolves anchors and chart/image relationships; `chartReferences.ts` supplies missing chart caches from internal worksheet ranges; `conditionalFormatting.ts` applies supported saved-value rules; `model.ts` defines cell/grid data. |
| `plaintext/` | The text front-end: `decode.ts` (BOM, UTF-8, UTF-16, Windows-1252 fallback), `classify.ts` (content-based classification and delimiter sniffing), `csv.ts` (quote-aware state machine), `index.ts` (single CSV worksheet or text pages), `paginate.ts` (text wrapping and pagination) and `metrics.ts` (the page geometry and text metrics the page renderers import as well). |
| `../formats.ts` | The format table: labels, extensions and MIME types, from which `SUPPORTED_EXTENSIONS`, `ACCEPT_ATTRIBUTE` and every "supported formats are …" message are generated. |
| `pptx/` | `index.ts` assembles a presentation: slide order from `sldIdLst`, layout/master resolution, inherited shape layers, background resolution, diagnostics. |
| `docx/` | The whole Word front-end: `styles.ts` (style chain, doc defaults, numbering), `text.ts` (paragraphs → run/picture/page-break blocks), `table.ts`, `drawing.ts` (inline and anchored pictures), `layout.ts` (pagination), `measure.ts` (the DOM-free height estimate) and `index.ts` (orchestration). |
| `pageContent.ts` | The data a self-drawn format returns instead of shapes, one kind per renderer. |
| `source.ts` | Normalises every accepted source (File, Blob, buffers, URLs) into bytes, with download limits, and maps any failure into the public error model. |
| `pdf.ts` | The pdf.js adapter: loads the engine lazily, reads page count, first-page geometry and metadata into a model, and returns the document as page content. Normalises pdf.js's own errors into the public error model. |

Parsing is synchronous after the download for the three OOXML formats and produces
one immutable model per source; the PDF path is asynchronous because pdf.js is
(and, unlike the OOXML parsers, it is allowed to be — the viewer already awaits the
load before committing it). Many unsupported constructs become diagnostics and, where the renderer can, a
labelled node. The compatibility matrix identifies features that are ignored
without diagnostics as well.

### Renderer (`src/renderer`, `src/components`)

| Module | Responsibility |
| --- | --- |
| `RenderShape.tsx` | Positioned box per shape: SVG geometry layer + HTML text layer, rotation/flips, groups, connectors, pictures with crop, unsupported placeholders. |
| `RenderText.tsx` | Paragraphs as absolutely placed blocks inside a flex column (vertical anchor), inline runs as spans, bullets as positioned spans. |
| `RenderTable.tsx` | Tables as absolutely positioned cells (merges produce the right spans). |
| `RenderChart.tsx` | Charts redrawn from parsed series data as inline SVG; exact values/hover details, labels, axes and legends. XLSX can synthesize series caches from internal ranges, leaving missing numeric points absent. Crowded category ticks are thinned without dropping points. |
| `style.ts` | Shared mapping of model values to CSS/SVG (fills, gradients, dash arrays, shadows, ids). |
| `pageRenderer.ts` | The seam a non-shape format implements: given a page, its natural size, the current scale and a `warn` callback, return a node. |
| `../components/pageRenderers.tsx` | Maps each `PageContent` kind to its page renderer; the only place a parser's output meets a component. |
| `PdfPage.tsx` | The PDF implementation of that seam: sizes a canvas from the scale × device pixel ratio (capped for memory), renders one page with pdf.js, cancels on unmount, and turns a failure into a labelled notice plus a diagnostic. |
| `TiffPage.tsx` | The TIFF implementation of that seam: imports the decoder on demand, decodes one directory to RGBA, draws it on a canvas, and reports a page it cannot decode. |
| `SheetPage.tsx` | Shared XLSX/CSV HTML grid: styling, merges, row/column headers, range selection/copy, saved scrolling and drawing overlays. XLSX uses stored heights and text spillover; CSV wraps at capped column widths with content-driven row heights. |
| `WorksheetTools.tsx`, `../utils/worksheet.ts` | Cross-sheet cell search (up to 500 matches), fit-to-width, clipboard copy with a two-second notice, and displayed-range CSV export. Export quotes fields and neutralizes formula-like text. |
| `measureParagraph.ts` | Measures laid-out text with a hidden mirror DOM (`kind: "layout"`), and falls back to the parser's metrics-based estimate (`kind: "estimate"`) where no DOM is available, for example in SSR or a test environment. `useDocument` injects it into the DOCX parser. |
| `DocumentViewer.tsx` | Public component: wires the hooks below together and owns the chrome (toolbar, thumbnails, diagnostics, fullscreen, presentation hint) and the imperative handle. Paginated formats use `SlideCanvas`; XLSX/CSV use `WorkbookView`. It maps the parsed model into `DocumentInfo`. |
| `WorkbookView.tsx` | The worksheet stage: measured scrolling viewport, sheet tabs, grid zoom buttons and `WorksheetTools`. Assigns worksheet panel IDs after mount for React 17 SSR compatibility. |
| `../hooks/` | `useDocument` (loading state machine, cancellation, media lifetime), `usePageNavigation` (controlled/uncontrolled page index), `useZoom` (fit modes, zoom ladder), `useRenderIssues` (buffered render diagnostics), `useViewerKeyboard` (focused and slideshow key handling), `usePresentationMode`, `usePresentationHint`, `useFullscreen`, `useElementSize`. |
| `SlideCanvas.tsx` | Scroll container, measurement, memoised slide view, per-slide error boundary — and the one line that delegates to a format's own page renderer. |
| `ViewerToolbar.tsx` | The chrome itself: a compact floating pill centred at the bottom of the stage, revealed by `:hover`/`:focus-within` in the stylesheet so it needs no visibility state of its own. Presentation only; every action and every piece of state belongs to `DocumentViewer`. |
| `ZoomMenu.tsx` | The zoom percentage and its dropdown of zoom stops and fit modes. The percentage and caret are one split button that opens on click or from the keyboard (never on hover), and it only calls back — it never owns the scale. |

Diagnostics reported during rendering (an undecodable image, for example) are
buffered and flushed in an effect after the commit, because a state update during
another component's render is not allowed by React.

## Security model

Every supported file is untrusted input, and the viewer is expected to run on
user-uploaded files:

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
7. **Third-party decoders**: pdf.js runs with `isEvalSupported: false`, so a PDF
   cannot reach `eval`, and the page rasters it produces are capped (3× CSS density,
   8 megapixels) so a hostile page size cannot exhaust canvas memory. A PDF is
   handed to pdf.js as a copy of the bytes, because pdf.js transfers the buffer to
   its worker. TIFF directories are walked with bounds checks and decoded into a
   canvas only when the page is visible; a directory that points outside the input,
   loops forever or declares an absurd size is refused rather than followed.
8. **Spreadsheet actions**: drawing assets and chart range references stay within
   the loaded archive. Formula text is not evaluated. CSV download and clipboard
   copying are user-initiated local actions; CSV export quotes fields and prefixes
   formula-like text with an apostrophe, while copied ranges preserve text.
9. **One format registry, one dispatch**: the format is decided from the bytes, so
   there is no path where a file name influences which parser runs.

## Performance model

- The archive is indexed once. PPTX themes, masters, layouts and charts are
  cached by part name; other formats have their own parsing paths.
- Only relevant parts are inflated — embedded workbooks, notes and diagrams are
  never decompressed.
- `pdfjs-dist` and the TIFF decoder are behind dynamic `import()`s, so they are
  downloaded only when a PDF or a TIFF is opened, and only when the consumer's
  bundler code-splits. Neither appears in the library bundle (`external` in the
  build).
- A `.docx` is paginated once, at load, by measuring each paragraph in a hidden
  mirror DOM in batches of up to 4,000 paragraphs; the measurement buffers are
  reused and removed afterwards.
- Media bytes are turned into URLs on first use, so a deck with 200 images only
  creates URLs for the slide being looked at (thumbnails are lazy-mounted by an
  `IntersectionObserver`).
- The visible slide is memoised; changing slides re-renders one slide.
- Slide geometry comes from the file. DOCX pagination measures text; worksheet
  layout measures its viewport and lets CSV row heights follow wrapped content.
- Worksheet rendering is capped, not virtualized: XLSX has a 10,000-cell grid
  budget, and CSV uses the same displayed-cell cap after parsing. Selection updates
  can rerender the active grid; inactive sheets are not mounted.

## Testing strategy

- **Fixtures are generated, not checked in**: `tests/fixtures/pptx.ts` writes real
  OOXML packages (masters, layouts, themes, slides, media, charts) with `fflate`,
  `tests/fixtures/docx.ts` writes WordprocessingML packages, `tests/fixtures/xlsx.ts`
  and `tests/fixtures/xlsxDashboard.ts` write SpreadsheetML workbooks with visual
  styles, chart references, drawings and conditional formatting, `tests/fixtures/tiff.ts` writes real IFD-based
  TIFFs, and `tests/fixtures/pdf.ts` assembles a real PDF byte by byte (objects,
  xref, trailer, metadata). No proprietary file is redistributed, and every edge case (missing
  relationship, missing part, empty slide, unsupported media, unsafe hyperlink,
  renamed format, encrypted file) can be produced exactly.
- **Core parser tests** do not depend on browser layout; visual worksheet
  regressions additionally exercise DOM output under jsdom.
- **Renderer and component tests** run under jsdom and exercise the component
  through the DOM the user sees: clicking toolbar buttons, pressing keys, reading
  the accessibility tree. PDF tests inject a fake pdf.js document through the
  `loadPdfJs` seam, which is why the adapter is typed structurally.
- **Browser tests** (`npm run test:browser`) run in real Chrome against the demo:
  they lay out a `.docx` into pages, rasterise a PDF page with pdf.js in a real Web
  Worker, draw a TIFF onto a canvas, and open a workbook grid — none of which a
  jsdom test can prove.
- **Security tests** assert the archive limits, ZIP-bomb rejection, path traversal
  filtering, XML entity refusal and hyperlink sanitisation.
- **Package tests** (`npm run test:package`) build, `npm pack`, then install the
  tarball into a React 19 ESM app (which renders a real deck and asserts the slide
  text reaches the DOM) and a React 17 CJS app (SSR), plus a consumer type-check
  against the published declarations.

## Known architectural limits

- The OOXML parsers are synchronous on the main thread; only PDF parsing can be
  moved off it, and only because pdf.js brings its own worker. A worker-based
  implementation for `.pptx`/`.docx`/`.xlsx` would be the obvious next step for very
  large files, at the cost of a bundling story for the worker.
- Word pagination is the viewer's own, driven by measured text. It cannot match
  Word exactly while floating shapes, headers/footers and missing fonts are out of
  the model; the divergences are reported rather than hidden.
- PDFs are rasterised, so their text is not selectable and not in the
  accessibility tree — the deliberate cost of reusing a real engine instead of
  re-implementing one.
- The style cascade is deliberately shallow: placeholder geometry and defaults are
  inherited, but PowerPoint's full style matrix (`fmtScheme` references, table
  styles) is approximated.
- Paginated-format thumbnails render the same page components as the main view.
  Visited thumbnails currently remain mounted; their working set is not bounded.
- Worksheet grids do not have virtualized rows or frozen panes. Search is limited
  to rendered cell data; omitted rows are neither searchable nor exported. Worksheet
  tool labels are currently English even when page-toolbar labels are localized.
