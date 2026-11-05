# PPTX and PDF performance review

Measured 1 November 2026 on this machine, installed Chrome, production demo at
`http://127.0.0.1:5274`, 1280 × 900 viewport, thumbnail sidebar enabled.

## Scope and current status

This is a historical measurement report from the PPTX/PDF review included in
0.2.1, not a fresh benchmark of 0.2.3. The worksheet tabs, XLSX styling/drawings,
CSV grid and tools added in 0.2.3 were not measured here. The recommendations
below remain unimplemented; added benchmarks are not runtime optimizations.
For current features and limits, see [README.md](../README.md) and
[COMPATIBILITY.md](../COMPATIBILITY.md). Re-run the commands below before using these
numbers for a new release, device or workload.

## Results

| File | Size | Runs | Median ready | Median first page mounted | Median PDF raster complete |
| --- | ---: | ---: | ---: | ---: | ---: |
| Project revenue deck, 30 slides | 1.04 MiB | 5 | 33 ms | 34 ms | — |
| Generated image-heavy PPTX, 120 slides | 20.82 MiB | 5 | 44 ms | 44 ms | — |
| Generated image-heavy PPTX, 600 slides | 104.08 MiB | 3 | 190 ms | 195 ms | — |
| Generated simple PDF, 600 pages | 1.39 MiB | 5 | 52 ms | 53 ms | 78 ms |

Timings begin at the file input change event, exclude network download and initial
app startup, and are medians of repeated local loads. Browser/HTTP caches may be
warm; these are not cold-start or mobile results. PPTX mounting does not include
completion of every image decode or browser compositing. PDF raster completion
means the main page's pdf.js render promise resolved, not all 600 pages rendered.
The generated PDF is simple text/vector content, not a scanned or image-heavy PDF.
No real user PDF was available in the repository.

Development results were slower: the project deck mounted in 88 ms, a generated
120-slide deck in 87 ms, and a generated 600-page PDF in 228 ms. Use production
measurements for optimization decisions; development React adds substantial work.

## Confirmed issues and recommended order

### 1. Bound the thumbnail working set

`src/components/ThumbnailSidebar.tsx` adds visible indices to a Set but never removes
them when they leave the observed area. Visited thumbnails remain mounted until
the viewer/document is replaced. It also creates a button for every slide/page.

Measured production working sets after browsing:

| File | Initially | After browsing |
| --- | --- | --- |
| Project deck | 9 rendered thumbnails; 1,859 DOM nodes | 30 thumbnails; 5,956 DOM nodes after 30 visits |
| 600-page PDF | 4 rendered thumbnails; 1,077,120 canvas pixels | 63 thumbnails; 13,787,136 pixels after 60 visits |

At four bytes per RGBA pixel, the PDF canvases alone grew from about 4.3 MB to
55.1 MB. These are backing-store estimates, not total process memory. Browser,
pdf.js, images, GPU resources, and DOM add further overhead.

Recommendation: virtualize thumbnail rows with a small overscan window and retain
only a bounded cache. Preserve active-thumbnail scrolling and keyboard focus.
Memoize PPTX thumbnail content so selecting a page does not rerender all retained
slide trees. Reset visibility/cache state on document changes.

### 2. Render PDF thumbnails at their displayed resolution

`src/components/PdfPage.tsx` clamps raster density to at least 0.5 for both main
pages and thumbnails. At this benchmark's 1× DPR, a 148 px thumbnail for an 816 px
natural page needs roughly 0.181× density, but receives 0.5×. It rasterizes about
7.6 times as many pixels as its displayed size requires.

Recommendation: distinguish thumbnail and main-page raster policies, using the
thumbnail's actual size × device pixel ratio. Keep the existing main-page density
and pixel ceilings. This would reduce this fixture's thumbnail pixel allocation
by roughly 87%, before any benefit from virtualization. Confirm visual quality at
1× and 2× DPR rather than lowering main-page quality globally.

### 3. Remove optional PDF work from the first-page critical path

`src/parser/pdf.ts` sequentially calls `getPage` for up to 200 pages to check page
sizes, then waits for optional metadata before returning the document model.
This happens before page 1 is displayed, even though only its size is required.

Recommendation: return after obtaining page 1 geometry, then validate other page
sizes and load metadata in bounded background work. Preserve mixed-size diagnostics,
cancellation, document disposal, and metadata callback semantics. This is a code
inspection finding; the isolated saving has not yet been measured.

### 4. Keep large PPTX parsing responsive

`src/parser/archive.ts` inflates ZIP data synchronously and
`src/parser/pptx/index.ts` parses all slides before displaying the first.
`src/parser/document.ts` invokes this work on the browser main thread.

The 104 MiB stress deck took 190 ms to reach ready on this machine. That total is
not a measurement of one uninterrupted task, but the synchronous parsing path can
block input, especially on slower devices.

Recommendation: move ZIP/XML parsing into a worker, transfer byte buffers, and
create media object URLs on the main thread. Consider lazy slide parsing afterward;
worker parsing is the more direct responsiveness improvement. Measure long tasks
on slower hardware before selecting a threshold for worker use.

### 5. Avoid repeated PDF rasterization during rapid zoom

The PDF rendering effect depends directly on `scale`; each scale change cancels
the current render and starts another. No bounded rendered-page bitmap cache is
present for revisiting pages.

Recommendation: update CSS zoom immediately, then schedule a high-quality raster
after a brief idle period. Add a small LRU cache keyed by document, page, and density,
with a total pixel budget. Do not introduce an unbounded canvas cache.

## Existing strengths

- PDF support is lazy-loaded; the demo already supplies a real pdf.js worker.
- Only the active PDF page and mounted thumbnails are rasterized.
- PDF canvases have an eight-million-pixel ceiling and rendering cancellation.
- PPTX relationships, themes, masters, layouts, and charts are cached during parsing.
- The main PPTX slide view is memoized.

## Benchmark improvements made during this review

- `BENCH_FILE` accepts an existing local file.
- `BENCH_BASE_URL` targets a production preview instead of the development server.
- Files over Playwright's 50 MiB buffer limit use temporary file uploads.
- `BENCH_BROWSE=1` reports thumbnail, DOM, and canvas growth after browsing.
- PDF canvases expose `data-render-state` to measure completed rasterization.
- The old `paint` output is now called `mounted` to avoid implying visible content
  was fully drawn when only its container existed.

Example commands, after building and starting a production preview:

```sh
npm run build:demo
npx vite preview --config examples/demo/vite.config.ts --host 127.0.0.1 --port 5274
# In another terminal:
BENCH_BASE_URL=http://127.0.0.1:5274 BENCH_FILE='Weekly Top Categories by Revenue — expanded brief.pptx' BENCH_RUNS=5 npm run bench
BENCH_BASE_URL=http://127.0.0.1:5274 BENCH_ONLY=report BENCH_RUNS=5 BENCH_BROWSE=1 npm run bench
BENCH_BASE_URL=http://127.0.0.1:5274 BENCH_ONLY=deck BENCH_SCALE=5 BENCH_MEDIA=heavy npm run bench
```

Validation: 232 existing unit/regression tests, TypeScript checks, production demo
build, and the actual-file, PDF, and large-PPTX benchmarks passed. The optimization
recommendations above have not been implemented or measured as before/after changes.
