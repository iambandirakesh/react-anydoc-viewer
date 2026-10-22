# Extending `react-anydoc-viewer` to the full upload matrix

**Question.** The upload allow-list in the screenshot accepts images, PDF, `.docx`,
`.xlsx`/`.xls`, `.pptx` and text/data (`.csv`, `.txt`, `.log`, `.json`, `.html`,
`.htm`) plus `.md`, at 20 MB per file and 10 files per upload. Can this viewer
open all of that, and what exactly has to change?

**Answer in one line.** Yes for everything except the legacy binary formats
(`.xls`, `.doc`, `.ppt`) — and the four seams the codebase already has
(bytes-based detection, "a page is a slide", the per-page renderer, the media
store) mean most of the new formats are *additive* rather than a rewrite. The
real engineering is Excel, because a spreadsheet is a scrollable grid and not a
paginated page.

This document is an analysis, not a change set. Every claim about the current
code below was read out of this repository; every claim about third-party
libraries is sourced in the last section.

### Status

| Phase | State |
| --- | --- |
| 0 — honest legacy/encrypted detection, format list from one table, exported extension list | **Implemented** (`src/parser/ole.ts`, `src/formats.ts`) |
| 1 — images | **Implemented** (`src/parser/image.ts`, `src/components/ImagePage.tsx`) |
| 2 — text and data (`.txt`, `.log`, `.json`, `.csv`, `.tsv`, `.html` as source) | **Implemented** (`src/parser/plaintext/`) |
| 3 — Markdown as a rendered document | **Outstanding** — currently shown as source text, with a diagnostic |
| 4 — Excel `.xlsx` | **Outstanding** |
| 5 — HTML in a sandboxed frame | **Implemented, then removed** (`htmlPreview`, `src/components/HtmlPage.tsx`): the maintainer decided the viewer should not render active content at all, so HTML is source-only, with no option |

Two things changed while implementing, and the sections below are left as they
were written so the reasoning is auditable:

- The text and CSV pages are drawn by their own page renderers (`TextPage`,
  `GridPage`) rather than by building shapes for the existing slide renderer. The
  `PageRenderer` seam already existed for exactly this (PDF uses it), exact
  monospace metrics matter more than model uniformity, and it removed any risk of
  disturbing the `.docx` layout pipeline. §5.1's "reuse the docx paginator" is
  therefore superseded.
- HTML went through both options in §5.3 and ended at the cautious one. The frame
  was built first: an empty `sandbox`, the `default-src 'none'` policy §5.3 called
  the real fix, never `allow-same-origin`, behind an opt-in `htmlPreview` prop, with
  a Chrome test proving a hostile page could neither script the host document nor
  get a response for its remote beacon. It was then removed on request — a viewer
  that promises never to execute file content is simpler to reason about than one
  that does it carefully — so §10's decision is **source only**, with no option and
  no frame, and the state of the art for a host that wants more is its own
  cookie-less origin.

---

## 1. Verdict per row of the matrix

| Category | Extensions | Verdict | How it would render | Size |
| --- | --- | --- | --- | --- |
| Images | `.png`, `.jpg`, `.jpeg`, `.gif`, `.webp` | **Yes — new** | One page per image, letterboxed; `<img>` over an object URL | S |
| PDF | `.pdf` | **Already supported** | pdf.js → canvas, via the existing page-renderer seam | — |
| Word | `.docx` | **Already supported** | Own parser → pages of shapes | — |
| Excel | `.xlsx` | **Yes — new** | Read-only windowed grid, one page (tab) per worksheet | L |
| Excel (legacy) | `.xls` | **No — report clearly** | OLE/CFB, BIFF8 binary; out of scope | S (message only) |
| PowerPoint | `.pptx` | **Already supported** | Own parser → pages of shapes | — |
| Text/data | `.txt`, `.log`, `.json` | **Yes — new** | Paginated monospace text (reuses the `.docx` paginator) | S |
| Text/data | `.csv` | **Yes — new** | Grid (same surface as `.xlsx`, no styles/dates) | M |
| Text/data | `.html`, `.htm` | **Yes — new, source only** | Paginated as source text like any other text file; never parsed, never framed, never executed | S |
| Markdown | `.md` | **Yes — new** | Markdown → block model → the `.docx` paginator | M |

Rows the screenshot *rejects* that are worth a deliberate decision because they
are nearly free once the above exists:

| Extension | Note |
| --- | --- |
| `.svg`, `.bmp` | `sniffImageKind()` already detects both; the image path gets them for nothing |
| `.avif`, `.ico`, `.tiff`, `.emf`, `.wmf` | `.avif` needs one more magic-byte branch; `.tiff`/`.emf`/`.wmf` are already known-undecodable and are reported, not guessed |
| `.xml`, `.yaml`, `.tsv` | Same "text page" path as `.txt`; `.tsv` is CSV with a fixed tab delimiter |
| `.xlsm` | Byte-identical container to `.xlsx` (`xl/workbook.xml`); the same parser opens it read-only. Server-side allow-list is a separate policy |
| `.zip` | `readArchive` already enforces a decompression-bomb guard, but there is nothing meaningful to *view*; keep refusing |
| `.doc`, `.ppt`, `.xls` | Genuinely unsupported: pre-2007 binary. See §7 for the honest-error fix |

Non-goals, stated plainly: Office fidelity, editing, formula evaluation, macro
inspection, font embedding, and printing fidelity. The library's contract is
"renders a documented subset faithfully and reports everything else".

---

## 2. What the current architecture already gives us

These four seams are why the answer is "mostly additive":

1. **Detection from bytes, never the name.** [`format.ts`](src/parser/format.ts)
   (`sniffContainer`, `probePackage`) plus
   [`document.ts`](src/parser/document.ts) (`detectFormat`,
   `parseDocumentBytes`) decide the format from magic bytes and OPC part names.
   A new format is a new branch, not a new code path through the components.
2. **"A page is a slide."** Every parsed document becomes a
   `PresentationModel` of `SlideModel`s ([`model/types.ts`](src/model/types.ts)).
   Navigation, thumbnails, zoom, fit modes, keyboard handling, presentation mode,
   the error boundary and diagnostics are all format-agnostic and come free.
   The `.docx` parser already proves this for a non-slide format: a page of a Word
   document is a slide whose shapes are paragraphs, tables and pictures.
3. **The per-page renderer seam.** [`pageRenderer.ts`](src/renderer/pageRenderer.ts)
   lets a format draw its own page when the content is not a shape list
   (PDF does exactly this, lazily importing `pdfjs-dist`). Spreadsheets, HTML and
   images plug in here.
4. **The media store + archive limits.** [`media.ts`](src/parser/media.ts)
   already turns part bytes into `<img>`-safe object URLs (and deliberately
   *not* into markup, so untrusted SVG cannot script the host page), and
   [`archive.ts`](src/parser/archive.ts) already enforces entry count, per-entry
   bytes, total bytes and a compression-ratio guard.

Consequences worth making explicit, because they set the shape of the work:

- **The library is one document at a time.** "10 files per upload" is a host-app
  concern; the viewer takes one `source`. Switching files disposes the previous
  media store correctly ([`usePresentation.ts`](src/hooks/usePresentation.ts)),
  so a host queue is a small wrapper, not a library feature.
- **20 MB per file maps onto existing knobs**, not new ones:
  `maxDownloadBytes` (default 128 MiB) and `limits.maxCompressedBytes` /
  `maxEntryBytes`. A host that mirrors the server's 20 MB cap sets both.
- **The "name is never trusted" principle collides with plain text.** `.txt`,
  `.log`, `.csv`, `.json`, `.md` and `.html` are the same bytes to a sniffer.
  See §6.1 — this is the one genuinely new *design* decision in the whole
  document.

---

## 3. Images — the cheap win

**Mechanism.** A new `image` container kind whose single page is one picture.

- Detection: reuse [`archive.ts#sniffImageKind()`](src/parser/archive.ts) — it
  already recognises png/jpeg/gif/bmp/webp/tiff/emf/wmf and text-SVG. Add one
  branch for AVIF (`ftyp` + `avif`/`avis` brand) and, if wanted, ICO.
- Page geometry: an image has no declared page size, so the page box is the
  image's natural pixel size converted to EMU (1 px = 9525 EMU, matching
  `EMU_PER_PX` in [`units.ts`](src/parser/units.ts)) — i.e. `fit="page"` fits it
  exactly and `fit="actual"` is pixel-accurate.
- Dimensions: read the header (PNG `IHDR`, JPEG SOF, GIF logical screen, WebP
  `VP8X`/`VP8`/`VP8L`, BMP info header) rather than waiting for a decode. This
  keeps parsing synchronous like the rest of the library and avoids a second
  async path; `decode` failure is still caught by the existing `ErrorBoundary`
  in [`SlideCanvas.tsx`](src/components/SlideCanvas.tsx).
- SVG stays an `<img src=objectURL>` — same rule the media store already
  documents. Never an iframe, never `dangerouslySetInnerHTML`.
- Big or malformed images: `limits.maxEntryBytes` already caps them; a
  multi-megapixel decode is bounded by the browser, but a `LIMIT_EXCEEDED`
  diagnostic for absurd dimensions is worth adding.
- Animated GIF/WebP just animate in `<img>`; EXIF orientation is applied by
  modern browsers (`image-orientation: from-image`), so no manual transpose.

**Effort: S.** New parser module + one detection branch + tests. No new
dependency, no new UI.

---

## 4. Excel `.xlsx` — the real project

This is the only item that needs a new *surface*, not just a new parser.

### 4.1 Parsing: hand-rolled, on the dependencies already in the tree

An `.xlsx` is an OPC package, so **`fflate` + `fast-xml-parser` — both already
direct dependencies — are sufficient.** That fits the project's dependency
story (the only heavyweight dependency today is `pdfjs-dist`, for PDF only) and
avoids the SheetJS distribution problem documented in §8.

Parts a read-only renderer needs:

| Part | What it provides |
| --- | --- |
| `xl/workbook.xml` | sheet list, sheet names, order, `r:id`s, `workbookPr/@date1904` |
| `xl/_rels/workbook.xml.rels` | `r:id` → worksheet part (never assume `sheet1.xml`) |
| `xl/worksheets/sheetN.xml` | rows, cells, `t=` types, cached `<v>`, `<f>`, `<mergeCells>`, `<cols>`, `<sheetFormatPr>`, `<dimension>` |
| `xl/sharedStrings.xml` | shared string table for `t="s"` indexes |
| `xl/styles.xml` | `cellXfs` → `numFmtId`, built-in and custom `numFmts` (dates) |
| `xl/media/*` | pictures anchored in sheets (optional; the archive filter must allow them) |

Two required changes to existing code:

- [`archive.ts`](src/parser/archive.ts) `USEFUL_PART` currently matches
  `[content_types].xml`, `*.rels`, `ppt/**`, `word/**` and `docProps/core.xml`
  — **nothing under `xl/` is ever inflated today**, so `xl/**` (minus
  `xl/vbaProject.bin`, embeddings, printer settings) has to be added. The
  existing per-entry and total limits already make that safe.
- [`format.ts`](src/parser/format.ts) `probePackage` already *recognises*
  `xl/workbook.xml` / `xl/workbook.bin` — it returns `format: null` with the
  label "an Excel workbook (.xlsx)". That becomes `format: "xlsx"`.
  `xl/workbook.bin` is `.xlsb` (binary sheets, unreadable as XML) and should
  keep returning `null` with a distinct label.

Why not a library: see §8 — the npm `xlsx` build is stale by the vendor's own
docs, `exceljs` is ~250 kB gzip for read-only viewing, and `read-excel-file`
(9.8 kB gzip, MIT) is a values-only parser that still needs the layout and
rendering work, i.e. it would save perhaps a day and add a dependency to a
project whose selling point is that it has almost none. If the goal were
"substitute for Excel", the calculus flips; for a viewer it does not.

### 4.2 The traps a naive implementation hits

Verified against the SpreadsheetML documentation:

1. **Relationship indirection** — worksheet file names are not derivable from
   sheet order; resolve `r:id` through the workbook rels.
2. **Namespace prefixes are arbitrary** — match local names, not `x:`.
3. **Sparse cells and rows** — empty rows/cells are omitted; place values by
   decoding the `r="C12"` reference, never by position.
4. **Five cell representations** — `t="s"` (shared string index), `t="inlineStr"`
   (`<is><t>`), `t="b"` boolean, `t="e"` error, `t="str"` formula string, plus
   the default numeric, and `t="d"` ISO-8601 dates in newer files.
5. **Dates are formatting, not a type.** A date is a number plus a number format
   resolved through `cellXfs/@numFmtId` (built-in ids 14–22, 45–47; anything
   ≥ 164 is custom). Also handle `date1904` and Excel's 1900 leap-year bug, and
   format times from the fractional part.
6. **Formulas have cached results, or none.** Display `<v>`; never evaluate
   `<f>`. A missing cache is a diagnostic, not a guess.
7. **Rich text** — shared/inline strings can be `<r>` runs with `<rPr>`; respect
   `xml:space="preserve"` or leading/trailing spaces vanish.
8. **Merged cells** — render the anchor text once, mark covered cells merged
   (the model's `TableCell.merged` exists for exactly this).
9. **Layout units** — column width is in "maximum digit width" characters and
   row height in points/twips; convert deliberately, and honour
   `defaultRowHeight`, `defaultColWidth`, hidden rows/columns and custom heights.
10. **Not every sheet is a grid** — chart sheets and dialog sheets exist and
    should be reported, not rendered as empty tables.

### 4.3 The surface: page = sheet, but a grid is not a page

A sheet of 40 000 rows cannot become shapes. Two parts make this tractable:

- **Pages are sheets.** `SlideModel.name` already exists, so sheet tabs map onto
  the thumbnail sidebar, the page counter and presentation mode for free. Most
  workbooks have a handful of sheets, so navigation stays sane.
- **The page body is a virtualised grid.** Render only the rows/columns inside
  the visible rectangle, driven by scroll position and `scale`. This needs a
  small extension to the existing seam in
  [`pageRenderer.ts`](src/renderer/pageRenderer.ts):
  - `PageRenderContext` currently carries `page`, `width`, `height`, `scale`,
    `interactive`, `media`, `warn`. Add the scroll viewport (offset + size) and a
    `mode: "page" | "thumbnail"` flag. Thumbnails are distinguished *today* only
    by `interactive: false` (the renderer's own comments call that "no links, no
    pointer interaction"), which is not enough information for a grid: rasterising
    a thousand-row sheet into a 148 px preview is pure waste, so a thumbnail should
    draw a bounded preview of the first N rows/columns instead. `mode` makes that
    distinction explicit rather than inferring it from a flag about links.
  - [`SlideCanvas.tsx`](src/components/SlideCanvas.tsx) already handles a page
    larger than the container (`radv-canvas-scroll`), so a wide sheet at
    `scale = 1` scrolls in both axes with no new CSS. It only needs to report the
    scroll offset and to pass `mode="page"`.
  - [`ThumbnailSidebar.tsx`](src/components/ThumbnailSidebar.tsx) passes
    `mode="thumbnail"`.
- **Default fit has to differ per format.** `fit="page"` on a 60-column sheet
  would render unreadable microtext. Options: (a) the format declares a
  suggested fit on `DocumentInfo` (e.g. `suggestedFit: "width"`, `maxInitialZoom: 1`),
  or (b) the viewer picks `width` for grid formats. (a) is the honest one: the
  host can override, and it generalises to HTML and text pages later.
- **Cell text rendering** reuses `TextBody`/`Run` and the existing text renderer,
  so alignment, bold, wrap and number-ish formatting work with no new code.
  Conditional formatting, cell borders/styles, charts, pivot tables and
  autofilters are reported, not implemented, in line with `COMPATIBILITY.md`.

**Effort: L** (parser + cells/styles/dates + windowed grid + thumbnails +
fixtures + docs). The parser alone is ~M; the surface is the rest.

**Test strategy.** Build fixtures in-process, as the repo already does for
`.pptx` and `.docx` ([`tests/fixtures/pptx.ts`](tests/fixtures/pptx.ts)): a ZIP
assembled with `fflate` at test time, covering shared strings, inline strings,
dates (1900 and 1904), merged cells, sparse coordinates, custom column widths and
a formula with a cached value. No binary fixture files, no network.

---

## 5. Text and data files

### 5.1 `.txt`, `.log`, `.json` — paginated text

Mechanism: decode, split into lines, measure in a monospace font, and paginate
with the same greedy algorithm `.docx` already uses
([`docx/layout.ts`](src/parser/docx/layout.ts)); the page size is US Letter or A4
(the `.docx` default in [`docx/index.ts`](src/parser/docx/index.ts) is already
US Letter, `12240 × 15840` twips). Each page's content is a handful of
`ShapeNode`s carrying a `TextBody`, so nothing in the renderer changes.

- **JSON**: pretty-print with a stable indent before paginating (a collapsible
  tree is a nicer future step, not a first step).
- **Encoding**: probe a BOM first (UTF-8, UTF-16LE/BE are the realistic ones),
  then try `TextDecoder("utf-8", { fatal: true })`, then fall back to
  `windows-1252` with a diagnostic. Modern `TextDecoder` also covers
  Shift_JIS and friends if a consumer needs them; the honest default is
  "UTF-8, else windows-1252, and say so".
- **Caps**: 20 MB of text is ~200k lines. Paginating all of it eagerly is the
  one real performance risk in this section; either cap the paginated prefix
  with a `LIMIT_EXCEEDED` diagnostic, or lay out pages lazily on demand (the
  nicer answer, and it composes with the windowing work needed for `.xlsx`).
- **Binary junk**: a non-text file that slips past detection must not produce
  10 000 pages of U+FFFD. The text sniffer's control-character ratio (§6.1)
  guards this, and a hard "looks binary" diagnostic is a good backstop.

**Effort: S**, mostly shared with the CSV/Markdown paths.

### 5.2 `.csv` — a grid without styles

Reuse the `.xlsx` grid surface; skip styles, dates and formulas.

- Parse with a **character-by-character state machine**, not `split(",")`:
  quoted fields, `""` escaping, embedded newlines and delimiters inside quotes
  (RFC 4180 is descriptive, not a guarantee).
- **Delimiter sniffing**: sample several records under `,`, `;`, tab and `|`,
  and prefer the candidate with a consistent field count across records. Locale
  Excel exports use `;` — counting commas or reading one line gets it wrong.
  Only count delimiters outside quotes.
- Detect the BOM before decoding; Excel-origin files can be UTF-16.
- Keep every value as text; never coerce. Inconsistent row lengths are a
  diagnostic rather than silently shifted columns.
- Formula injection (`=`/`+`/`-`/`@` prefixes) is an *export* hazard, not a
  viewing one — the viewer never hands the text back to a spreadsheet — but it
  belongs in `SECURITY.md` as a note for hosts that re-export.

**Effort: M** (the sharing with `.xlsx` is what keeps it from being L).

### 5.3 `.html`, `.htm` — the only genuinely dangerous format

An uploaded HTML file is active content. Anything that renders it must assume it
is hostile.

Recommended: **render in a sandboxed iframe** — `sandbox=""` (no
`allow-scripts`, no `allow-same-origin`, no popups, no top-navigation) with
`srcdoc`, plus a restrictive CSP and `referrerpolicy="no-referrer"`. That blocks
scripts, forms, popups and top navigation, and gives the document an opaque
origin. Caveats to write down rather than gloss over:

- Sandboxing does **not** stop the document's own network loads (a remote image
  is still a beacon). Blocking `img-src` via CSP, or rewriting resource URLs, is
  the real fix — and it is the same problem the `.docx`/`.pptx` parsers solve by
  only ever reading bytes out of the archive.
- `srcdoc` inherits the embedding document's URL as its base, so relative URLs
  resolve against the host app. A CSP that denies everything external is
  therefore not optional.
- The strongest isolation is a separate, cookie-less origin — that is a *host
  application* responsibility, and belongs in the docs as a deployment note.
- Alternative for the cautious: show the HTML **as source text** (the `.txt`
  path) and offer a "render preview" toggle. Zero risk, near-zero cost.

**Effort: M**, mostly security review and tests, not code. Worth deciding with
the user: sandboxed render, source-only, or both.

### 5.4 `.md` — Markdown without an HTML sink

The tempting route is `marked` (13.5 kB gzip, MIT) → HTML string →
`dangerouslySetInnerHTML`. That introduces an XSS surface and needs DOMPurify
(~11 kB gzip) on top.

The route that fits this codebase better: **convert Markdown into the block
model the `.docx` parser already produces** (paragraph, heading, list, table,
code block, image) and let the existing paginator and text renderer draw it.
No HTML string, no `dangerouslySetInnerHTML`, no sanitizer, and headings/lists/
tables/code get the Word pipeline's behaviour for free — including
`sanitizeUrl()` ([`utils/sanitize.ts`](src/utils/sanitize.ts)) for links, which
already exists and already reports blocked schemes.

A hand-written subset (headings, paragraphs, emphasis/strong/inline code, fenced
code, blockquotes, ordered/unordered lists, thematic breaks, pipe tables, links,
images-as-labelled-placeholders, GFM task lists) is realistic. If full CommonMark
conformance is ever required, `marked` can be swapped in *behind the same block
model*, so that decision stays reversible.

**Effort: M** (subset parser + block mapping + tests).

---

## 6. Cross-cutting work

### 6.1 The one new design decision: sniffing plain text

The library's rule is "the bytes decide, never the file name". That rule works
because the formats are containers with signatures. Six of the new text types
have no signature at all.

Proposal, in priority order, in `sniffContainer`/`detectFormat`:

1. Binary signatures first: `%PDF-`, OLE/CFB, ZIP/OPC (unchanged).
2. Image magic bytes (`sniffImageKind`).
3. Only then a **text heuristic**: no NUL bytes in the first few KiB, a low
   ratio of control characters, and successful decode — then sub-classify by
   content, not by name: valid JSON; an HTML doctype or `<html`/`<body` tag;
   Markdown heading/fence/list markers; CSV delimiter consistency (≥ 2 rows and
   a uniform field count); otherwise plain text.
4. Anything else: the current "unknown" refusal.

The heuristic is a judgement call, so give it an escape hatch: an optional
`contentHint?: "text" | "csv" | "markdown" | "html" | "json"` prop, documented
as "the bytes still decide *whether* we can read it; the hint only decides how a
plain-text file is presented". That keeps the principle intact and makes the
ambiguous cases the host's call.

### 6.2 Public API and error model

- [`errors.ts`](src/errors.ts): widen `DocumentFormat` (e.g. add `"png"`… no —
  prefer `"image" | "xlsx" | "csv" | "text" | "markdown" | "html"`) and
  `SUPPORTED_FORMATS`. **This is a breaking change for consumers with exhaustive
  `switch` statements**, so the release that ships it is `0.2.0`, not `0.1.2`.
- Error messages that enumerate formats (`unsupportedFormatMessage` in
  `format.ts`, several strings in `archive.ts`, `README.md`, `COMPATIBILITY.md`)
  should be generated from `SUPPORTED_FORMATS` so they cannot drift again.
- New public helpers worth exporting: `SUPPORTED_EXTENSIONS` and a ready-made
  `ACCEPT_ATTRIBUTE` string, so a host's `<input accept>` and file picker cannot
  disagree with the viewer — and can be kept in step with the *server's*
  allow-list from the screenshot.
- `DocumentInfo`: optionally carry a page noun ("sheet", "page", "slide") and the
  suggested fit from §4.3, so labels and initial zoom stop being slide-centric.

### 6.3 Rendering seams

- `PageRenderContext`: add `mode: "page" | "thumbnail"` and the scroll viewport
  (this is the enabling change for `.xlsx`/`.csv` and for cheap thumbnails).
- New renderers: `GridPage` (windowed, reused by `.xlsx` and `.csv`), and reuse of
  the existing text renderer for everything else. (`HtmlPage`, the sandboxed-frame
  renderer, was built and then removed — see the status table above.)
- [`SlideCanvas.tsx`](src/components/SlideCanvas.tsx) and
  [`ThumbnailSidebar.tsx`](src/components/ThumbnailSidebar.tsx): pass the new
  context fields; no other behavioural change.
- [`DocumentViewer.tsx`](src/components/DocumentViewer.tsx): apply the suggested
  fit, and expose `contentHint`.

### 6.4 Lazy loading, bundle budget

`pdfjs-dist` is imported only when a PDF is opened, and the comments in
[`pdf.ts`](src/parser/pdf.ts) explain why. New parsers must follow the same
pattern (`await import(...)` inside the dispatch for that format) so `.pptx`,
`.docx` and `.pdf` users never pay for the new code. The package gate already
inspects the built output
([`scripts/verify-package.mjs`](scripts/verify-package.mjs)); it should grow an
assertion that the new features are separate chunks and that the base bundle has
not regressed.

### 6.5 Tests, fixtures, demo and docs

| Area | Work |
| --- | --- |
| Unit tests | New suites for spreadsheet (cells/styles/dates/merges), CSV state machine + delimiter sniffing, text pagination + encodings, Markdown → blocks, image headers |
| Fixtures | Build `.xlsx`/`.csv`/`.md`/text fixtures in-process with `fflate`, mirroring the two existing builders [`tests/fixtures/pptx.ts`](tests/fixtures/pptx.ts) and [`tests/fixtures/docx.ts`](tests/fixtures/docx.ts) |
| Security tests | Extend [`tests/security/security.test.ts`](tests/security/security.test.ts) (it already builds hostile packages with `fflate` and asserts parts like `docProps/thumbnail.jpeg` are never inflated) with the new surfaces: oversized/hostile spreadsheets, decompression ratio, HTML that must not execute, SVG that must not script |
| Browser tests | Extend [`tests/e2e/viewer.spec.ts`](tests/e2e/viewer.spec.ts): a workbook opens, tabs navigate, the grid scrolls, markdown renders, an HTML file does not execute script (assert no alert/global side effect), an image fits |
| Package gate | Extend `verify-package.mjs` and [`scripts/prepublish-guard.mjs`](scripts/prepublish-guard.mjs) with the new exported types and chunks |
| Demo | Sample workbook/text/markdown files in [`examples/demo/src`](examples/demo/src) and the picker's `accept` list |
| Docs | `README.md` format table and the `accept` snippet, a new `COMPATIBILITY.md` section per format, `ARCHITECTURE.md` seams, `SECURITY.md` (untrusted markup policy, HTML origin advice, CSV export note), `CHANGELOG.md` |

### 6.5.1 Existing files that will change, and existing tests that will start failing

Worth knowing up front, because these are assertions of *today's* behaviour:

- [`tests/parser/format.test.ts`](tests/parser/format.test.ts) asserts
  `sniffContainer("<html></html>")` is `"unknown"`, and expects a package with
  `xl/workbook.xml` to be refused. Both expectations invert once text sniffing
  and `.xlsx` land.
- [`tests/parser/docx.test.ts`](tests/parser/docx.test.ts) builds an `.xlsx`
  package to assert it is rejected — same inversion.
- [`examples/demo/src/App.tsx`](examples/demo/src/App.tsx) has a hand-written
  `ACCEPT` list and a `FORMAT_LABELS` map (`Record<string, string>`, so it will
  not fail to compile — it will silently fall back to a generic label). Both
  should come from the library once the extension list is exported, which is
  precisely the drift this analysis proposes to remove.
- [`scripts/verify-package.mjs`](scripts/verify-package.mjs) type-checks an
  example consumer against the built declarations and builds a synthetic `.pptx`
  with `fflate`; new formats need their own consumer snippets and fixtures there,
  and the base bundle assertion from §6.4.
- [`scripts/prepublish-guard.mjs`](scripts/prepublish-guard.mjs) will need the new
  extension list and any new public exports.

### 6.6 The 20 MB / 10 files limits

| Screenshot limit | Where it belongs |
| --- | --- |
| 20 MB per file | Host sets `maxDownloadBytes={20 * 1024 * 1024}` and `limits={{ maxCompressedBytes, maxEntryBytes }}`; the viewer already refuses with `LIMIT_EXCEEDED` |
| 10 files per upload | Host-side queue; the viewer takes one source. Document the recipe: keep a file list, switch `source`, let `usePresentation` dispose the old media store |
| "Any other file type is rejected" | Server policy, but mirror it client-side with the exported extension list/`accept` attribute so the two never disagree |

---

## 7. Fixing an existing inaccuracy: legacy `.doc`/`.ppt`/`.xls`

Today an OLE/CFB container is reported as `ENCRYPTED` with the message "This looks
like an encrypted or legacy binary (.doc, .ppt, .xls) Office file", and
`COMPATIBILITY.md` documents it that way. For a genuine `.xls` — which the
screenshot explicitly allows — that error is misleading.

An OLE compound file stores its streams in a directory whose names are discoverable:
an encrypted OOXML file contains `EncryptedPackage`/`EncryptionInfo`, whereas a
legacy `.xls` contains `Workbook`/`Book`, a `.doc` contains `WordDocument`, and a
`.ppt` contains `PowerPoint Document`. Even a bounded UTF-16 name scan over the
first sectors distinguishes the two cases, which would let the viewer say
"this is a legacy Excel 97–2003 workbook, which the viewer cannot display" and
add `UNSUPPORTED_FORMAT` for the true legacy case instead of mislabelling it as
encrypted. Small, self-contained, and it removes a documented lie. **Effort: S.**

---

## 8. Why hand-rolled over third-party parsers

Researched 2026 facts (sources at the end):

| Option | License | Size (min+gzip) | Legacy `.xls` | Verdict here |
| --- | --- | --- | --- | --- |
| `xlsx` (SheetJS CE, public npm) | Apache-2.0 | ~136 kB (0.18.5, published 2022) | Yes (full build) | npm build is stale by the vendor's own docs; current release (0.20.3) is CDN-only, vendored by hand — a supply-chain and bundling wart |
| `@e965/xlsx` (npm fork) | Apache-2.0 | ~136 kB | Yes | Third-party republish; last version evidence is old — needs independent review |
| `exceljs` | MIT | ~250 kB | No | Read/write workbook library; far too much for read-only viewing |
| `read-excel-file` | MIT | ~9.8 kB | No | Smallest credible parser, but values-only: layout, grid and rendering are still our work, so it saves ~a day and adds a dependency |
| `marked` | MIT | ~13.5 kB | — | Outputs HTML, no sanitising → needs DOMPurify (~11 kB) and an HTML sink; the block-model route avoids both |
| `markdown-it` | MIT | ~39.6 kB | — | Same HTML-sink issue, larger |
| `DOMPurify` | Apache-2.0/MPL-2.0 | ~11.2 kB | — | Only needed if an HTML *string* is ever injected; the block-model route avoids it. Needed conceptually for `.html`, where the iframe does the isolating |
| Luckysheet / x-spreadsheet | MIT | large (+ parser) | No | Full editors; Luckysheet is unmaintained upstream and points at Univer. Wrong tool for a viewer |

`fflate` and `fast-xml-parser` already in `dependencies` make a read-only `.xlsx`
and a Markdown subset achievable with **zero new runtime dependencies** — which
is the strongest version of this library's pitch ("no paid API, no SDK licence,
no external service", and now: no new dependency for anything but PDF).

The one place a dependency is genuinely attractive: full CommonMark conformance.
If that is a requirement, `marked` behind the block model is the smaller risk.

---

## 9. Phasing

Each phase is independently shippable and verifiable with the existing gates
(`npm run typecheck`, `npm test`, `npm run test:package`, `npm run test:browser`).

| Phase | Content | Effort | Ships as |
| --- | --- | --- | --- |
| 0 | Honest legacy/OLE detection (§7); format lists generated from one table; export extension list + `accept` attribute | S | ✅ done |
| 1 | Images (incl. `.svg`, `.bmp`, `.avif`) | S | ✅ done |
| 2 | Text/data: `.txt`, `.log`, `.json`, `.csv`, `.tsv`, `.html`-as-source + the text sniffer and `contentHint` | M | ✅ done |
| 3 | Markdown rendered as a document (blocks → pages) | M | next |
| 4 | Excel `.xlsx`: read-only parser, windowed grid, page-per-sheet, thumbnails, per-format fit | L | next |
| 5 | HTML in a sandboxed frame behind a toggle, with CSP and a browser test proving no execution | M | ✅ done, then removed on request — HTML is source-only |
| — | Legacy `.doc`, `.ppt`, `.xls`, `.xlsb`, plus `.zip` | (non-goal) | reported clearly, never rendered |

Acceptance criteria that must hold at every phase:

1. Format detection is still from bytes; a renamed or extension-less file loads.
2. A page is still a slide — navigation, thumbnails, zoom, presentation mode,
   keyboard handling and the imperative handle keep working unchanged.
3. Unsupported content is *reported* (diagnostic + labelled placeholder), never
   silently dropped.
4. Existing `.pptx`/`.docx`/`.pdf` bundle size does not regress (new code is
   lazy-loaded).
5. `npm run test:package` still validates the published bundle, including the new
   chunks and exports.

---

## 10. Decisions this analysis needs from the user

1. **HTML**: render in a sandboxed frame, show source only, or both behind a
   toggle? (Security posture vs. usefulness.) — **answered, finally: source only.**
   Both was implemented first (`htmlPreview`), then removed: a viewer that cannot
   execute file content at all is a smaller promise to keep.
2. **Excel ambition**: read-only values + essential layout (recommended), or an
   editor-grade grid (different project, ~10× the work, and Luckysheet/Univer
   territory).
3. **Legacy `.xls`/`.doc`/`.ppt`**: leave unsupported (recommended, with the
   honest error from §7) or adopt SheetJS-and-friends for BIFF8?
4. **Markdown**: hand-written subset (recommended, no XSS sink, no dependency) or
   `marked` + sanitizer for full CommonMark?
5. **Multi-file uploads**: host-side queue only (recommended), or a file-list
   feature in the viewer?

---

## Sources

Empty-safe parsing, dates, layout and library options:

- SheetJS installation/registry status — <https://docs.sheetjs.com/docs/getting-started/installation/nodejs/>
- `xlsx` on npm (0.18.5, 2022) — <https://www.npmjs.com/package/xlsx>
- SheetJS CDN and licensing — <https://cdn.sheetjs.com/>
- SheetJS browser build formats (XLS support only in the full build) — <https://docs.sheetjs.com/docs/getting-started/installation/standalone>
- `@e965/xlsx` fork — <https://www.npmjs.com/package/@e965/xlsx> · discussion: <https://github.com/SheetJS/sheetjs/issues/2825>
- ExcelJS — <https://github.com/exceljs/exceljs> · size <https://bundlephobia.com/package/exceljs>
- `read-excel-file` — <https://github.com/catamphetamine/read-excel-file> · size <https://bundlephobia.com/package/read-excel-file>
- SpreadsheetML structure, shared strings, formulas — <https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/structure-of-a-spreadsheetml-document> · <https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-shared-string-table> · <https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas>
- Luckysheet status — <https://github.com/dream-num/Luckysheet> · Luckyexcel — <https://github.com/dream-num/Luckyexcel>
- x-spreadsheet — <https://github.com/myliang/x-spreadsheet>

Markdown, sanitising, iframes, CSV and encodings:

- marked — <https://github.com/markedjs/marked> · markdown-it — <https://github.com/markdown-it/markdown-it> · micromark — <https://github.com/micromark/micromark> · remark — <https://github.com/remarkjs/remark>
- Sizes — <https://bundlephobia.com/package/marked> · <https://bundlephobia.com/package/markdown-it> · <https://bundlephobia.com/package/micromark> · <https://bundlephobia.com/package/remark>
- DOMPurify — <https://github.com/cure53/DOMPurify> · bypass classes — <https://github.com/cure53/DOMPurify/wiki/Attack-Classes-&-Bypass-History> · size <https://bundlephobia.com/package/dompurify>
- rehype-sanitize — <https://github.com/rehypejs/rehype-sanitize>
- MDN iframe sandbox and caveats — <https://developer.mozilla.org/en-US/docs/Web/HTML/Reference/Elements/iframe>
- RFC 4180 — <https://www.rfc-editor.org/info/rfc4180/> · Papa Parse (delimiter detection, formula-injection note) — <https://www.papaparse.com/docs>
- TextDecoder and the Encoding Standard — <https://developer.mozilla.org/en-US/docs/Web/API/TextDecoder> · <https://developer.mozilla.org/en-US/docs/Web/API/Encoding_API/Encodings>
