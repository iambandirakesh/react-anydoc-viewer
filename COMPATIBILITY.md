# Compatibility matrix

What `react-anydoc-viewer` renders today, what it renders approximately, and what it
deliberately reports as unsupported, for every format it opens: PowerPoint
(`.pptx`), Word (`.docx`), Excel (`.xlsx`, `.xlsm`), PDF, images (including TIFF),
delimited text (`.csv`, `.tsv`) and text files (`.txt`, `.log`, `.json`, `.html`).

Legend:

- **Implemented** — rendered from the file; colour, geometry and text styling come
  from the presentation itself.
- **Partial** — rendered, but with a documented simplification.
- **Reported** — not rendered; a labelled placeholder and a diagnostic appear
  instead of silently dropping the element.
- **Not supported** — ignored, with or without a diagnostic (noted below).

Everything in the "Reported" column produces an entry in the diagnostics panel
(`debug`) and in `onDiagnostics`, with a stable code such as
`geometry.preset-unsupported` or `ole.unsupported`.

## Presentation package

| Feature | Status | Notes |
| --- | --- | --- |
| Format detection from bytes | Implemented | `%PDF-` header, OLE/CFB signature, ZIP/OPC signature; the extension and MIME type are never trusted |
| OPC/ZIP package reading (`[Content_Types].xml`, `_rels/.rels`) | Implemented | Office-document part is discovered through relationships, not a hard-coded path |
| Package kind from part names (`ppt/presentation.xml`, `word/document.xml`, `xl/workbook.xml`) | Implemented | A renamed or extension-less package still opens correctly |
| Other OPC packages (`.xlsb`, OpenDocument) | Reported | `UNSUPPORTED_FORMAT` naming what the package actually is |
| Slide order via `p:sldIdLst` + `r:id` | Implemented | File names and numeric suffixes are never used for ordering |
| Slide size (`p:sldSz` cx/cy) and aspect ratio | Implemented | Slide is never stretched; `fit` scales it uniformly |
| Slide master, layout and theme resolution | Implemented | Layout → master → theme chain per slide |
| Multiple masters/layouts | Partial | All layouts of the first master are read; a deck with several masters uses the first |
| Core properties (title) | Implemented | Exposed on `onLoad` |
| Notes slides, comments, custom XML, embedded workbooks | Not supported | Never inflated (they are not in the useful-parts filter) |
| Encrypted / password-protected packages | Reported | `ENCRYPTED` error, recognised from the `EncryptedPackage`/`EncryptionInfo` streams in the OLE/CFB container |
| Legacy binary `.ppt` / `.doc` / `.xls` | Reported | `UNSUPPORTED_FORMAT`, naming the format (e.g. "a legacy Excel 97-2003 workbook (.xls)") and suggesting an export. Identified from the compound file's stream names (`WordDocument`, `PowerPoint Document`, `Workbook`/`Book`) |
| An OLE/CFB container whose streams cannot be read | Reported | `ENCRYPTED`, the conservative reading: it is either an encrypted package or a binary format |
| Corrupted/truncated package | Reported | `CORRUPT_PACKAGE` error, or per-part diagnostics |

## Slide content

| Feature | Status | Notes |
| --- | --- | --- |
| Shape tree (`p:sp`, `p:pic`, `p:graphicFrame`, `p:grpSp`, `p:cxnSp`) | Implemented | Connectors render as lines with arrow heads |
| Z-order (document order) | Implemented | Inherited layout/master shapes render behind slide shapes |
| Transforms: `off`, `ext`, `rot`, `flipH`, `flipV` | Implemented | Creates the CSS transform for each shape |
| Group coordinate spaces (`chOff`/`chExt`) | Implemented | Nested groups map through parent transforms |
| Placeholder inheritance (position, size, text styles) | Partial | Placeholders inherit geometry and defaults through master/layout; the full style cascade is not simulated |
| Slide backgrounds: solid, gradient, picture | Implemented | `p:bg` on the slide, layout or master |
| Hidden shapes (`hidden="1"`) and `show="0"` slides | Implemented | Skipped, as PowerPoint does |
| `mc:AlternateContent` | Partial | The fallback branch is used (conservative) |
| SmartArt (`dgm:`) | Reported | Labelled placeholder |
| Embedded OLE objects | Reported | Labelled placeholder |
| Video/audio placements | Reported | Labelled placeholder |
| Unknown graphic frames | Reported | Labelled placeholder naming the frame type |

## Text

| Feature | Status | Notes |
| --- | --- | --- |
| Runs, paragraphs, multiple levels | Implemented | |
| Font family, size, bold, italic, underline, strikethrough | Implemented | Theme font scheme (`+mj-lt`, `+mn-lt`) resolved |
| Run colour, highlight, theme colours, colour modifiers | Implemented | `srgbClr`, `schemeClr` with `lumMod`/`lumOff`/`tint`/`shade`/`alpha`/`satMod` |
| Superscript/subscript (`baseline`) | Implemented | |
| Alignment, indent/margin, hanging indents | Implemented | |
| Line spacing (`spcPct`, `spcPts`), space before/after | Implemented | |
| Bullets: `buChar`, `buAutoNum` (all types), `buNone`, `buFont`, `buClr` | Implemented | Numbering counters restart per level group |
| Vertical anchor and text insets | Implemented | |
| RTL paragraphs (`rtl="1"`) | Implemented | `direction: rtl` on the paragraph |
| Word wrapping and overflow | Partial | Browser line breaking; overflow is visible rather than clipped exactly like PowerPoint |
| Auto-fit (`normAutofit`/`spAutoFit`) | Partial | Recorded font scale and line-spacing reduction are applied; PowerPoint's fitting algorithm is not re-run |
| Fonts | Partial | Only fonts the browser can resolve; missing fonts fall back to the system stack |
| Text-on-a-path / WordArt | Reported | Rendered as a plain text box |
| Animations and transitions | Not supported | Not parsed (no diagnostics; they do not affect a static view) |

## Fills, lines and effects

| Feature | Status | Notes |
| --- | --- | --- |
| Solid, gradient (linear/radial), picture and patterned fills | Partial | Patterns approximate to the foreground colour |
| `noFill`, `grpFill` (inherit from group) | Implemented | |
| Picture fill cropping (`srcRect`) and tiling (`tile`) | Implemented | Non-uniform scaling preserved |
| Outlines: colour, width, dash (`prstDash`), cap, join, arrows | Implemented | |
| Outer shadow (`outerShdw`) | Implemented | Mapped to a CSS drop-shadow filter |
| Inner shadow | Partial | Drawn as an outer drop shadow; `effect.inner-shadow-approximated` |
| Glow, soft edge, reflection, preset shadow, fill overlay, 3-D bevel | Reported | Diagnostic `effect.unsupported` |
| Transparency (`alpha`) on fills and lines | Implemented | |
| Theme fill/line/effect styles (`style` element + `fmtScheme`) | Partial | `lnRef`/`fillRef`/`effectRef` are resolved for common cases; complex style matrices fall back to the theme default |

## Geometry

| Feature | Status | Notes |
| --- | --- | --- |
| 108 named preset geometries (`prstGeom`) | Implemented | Drawn as real SVG paths, with adjustment values honoured |
| Other preset names | Reported | `geometry.preset-unsupported`, drawn as a rectangle |
| Custom geometry (`custGeom`: `pathLst`, `moveTo`, `lnTo`, `cubicBezTo`, `quadBezTo`, `arcTo`, `close`) | Implemented | |
| Rotation/flipping of groups and shapes | Implemented | |

## Images

| Feature | Status | Notes |
| --- | --- | --- |
| PNG, JPEG, GIF, WebP, BMP, SVG | Implemented | Sniffed from magic bytes; declared content type is not trusted |
| Cropping (`srcRect`) and aspect preservation | Implemented | |
| Duplicate/reused images | Implemented | One object URL per media part, revoked on unmount |
| TIFF | Implemented | Decoded by the viewer's lazy TIFF reader and drawn on a canvas |
| EMF/WMF | Reported | Browsers cannot decode them; labelled placeholder |
| Very large images | Implemented | Decoded lazily by the browser; archive limits protect the parser |
| Image decode failure at render time | Reported | `picture.decode` diagnostic and a labelled placeholder; the rest of the slide still renders |

## Tables

| Feature | Status | Notes |
| --- | --- | --- |
| Rows, columns, explicit widths and heights | Implemented | |
| Cell text, fills, text anchor, margins, vertical anchor | Implemented | |
| `gridSpan`, `rowSpan`, `hMerge`, `vMerge` | Implemented | |
| Per-edge borders | Implemented | |
| `tableStyles.xml` banding and theme table styles | Partial | `table.style-approximated` diagnostic; a neutral grid is used |
| Nested tables | Reported | Not rendered inside cells |

## Charts

| Feature | Status | Notes |
| --- | --- | --- |
| Bar/column (clustered, stacked), line, area, pie, doughnut, scatter | Implemented | Drawn from cached plot data in the chart part |
| Axis titles, category and value axes, gridlines, tick labels | Implemented | |
| Legend (position, overlay) and value data labels | Implemented | |
| Series colours from `spPr`, falling back to the theme accent palette | Implemented | |
| `bubbleChart`, `radarChart`, `stockChart`, `surfaceChart` | Reported | Labelled placeholder |
| 3-D chart variants (`bar3DChart`, `pie3DChart`, …) | Partial | Rendered as the 2-D equivalent |
| Charts without cached values | Reported | `chart.empty` / `chart.unreadable`; a live chart is never recalculated |
| Chart types with a chart part missing | Reported | `chart.no-relationship` / `chart.bad-relationship` |

## Links and interaction

| Feature | Status | Notes |
| --- | --- | --- |
| Hyperlinks on runs and shapes | Implemented | `http`, `https`, `mailto`, `tel` only |
| `javascript:`, `data:`, `vbscript:` and other unsafe schemes | Reported | `hyperlink.blocked`; rendered as plain text, never as a link |
| Links disabled in non-interactive renders | Implemented | `interactive={false}` renders text without anchors |
| Action buttons / macros / slide-jump actions | Reported | `hyperlink.action` / `hyperlink.internal`; not executed and not followed |
| Slideshow animations, transitions, presenter view | Not supported | Out of scope for a viewer |

## Word documents (`.docx`)

The document is parsed from WordprocessingML and then **paginated by the viewer**,
so a page of a `.docx` is a slide as far as everything else in the library is
concerned: navigation, thumbnails, zoom, fit modes and presentation mode are the
same code paths.

| Feature | Status | Notes |
| --- | --- | --- |
| Page size (`w:pgSz`), orientation and margins (`w:pgMar`) | Implemented | Twips → EMU → CSS px; a landscape section's dimensions are normalised |
| Sections | Partial | Section breaks become page breaks; the last section's page setup is applied to the whole document (`docx.section-break`) |
| `w:body` block order (paragraphs, tables) | Implemented | Content controls, `w:ins` and custom XML wrappers are unwrapped; `w:del` and `w:moveFrom` are skipped |
| Paragraph properties: alignment, indents, spacing before/after, line spacing, keep-with-next | Implemented | |
| Paragraph borders and paragraph shading (`pBdr`, `pPr/shd`) | Reported | `docx.paragraph.borders` / `docx.paragraph.shading`; omitted |
| Run properties: bold, italic, underline, strike, `vertAlign`, size, font, colour, highlight, caps, `w:sym`, `w:br`, `w:noBreakHyphen` | Implemented | Half-points and twips converted; `w:rFonts` names used as written |
| Named styles, `basedOn` inheritance, doc defaults, built-in heading and list styles | Implemented | Resolved through the style chain, including latent/built-in fallbacks |
| Lists from `numbering.xml`: bullet, decimal, roman, alpha, `lvlText`, indents, restarts | Implemented | Counters advance per level group |
| Tables: grid, `gridSpan` merges, `rowSpan`/`vMerge`, cell shading, borders, row heights | Implemented | Nested tables are flattened with `docx.table.nested`; images inside cells are skipped (`docx.table.image`) |
| Inline pictures from `a:blip`/`r:embed` | Implemented | Scaled to the drawing's extent |
| Hyperlinks | Implemented | Same scheme allowlist as slides |
| Explicit page breaks and keep-with-next | Implemented | |
| Page breaks | Partial | Computed from measured paragraph and line heights, so a break can differ from Word by a line (`docx.layout.estimated` when measurement is unavailable, `docx.page.overflow` when a block exceeds a page) |
| Headers and footers | Reported | `docx.header-footer`; not rendered |
| Footnotes and endnotes | Reported | `docx.notes` |
| Comments | Reported | `docx.comments` |
| Equations (`m:oMath`) | Reported | `docx.equation`; skipped |
| Floating (anchored) pictures | Partial | Drawn at their page offset; text wrapping is not simulated (`docx.picture.wrap`) and text-relative anchors fall back to the flow (`docx.picture.alignment`) |
| Text boxes, `wps` drawing content | Reported | `docx.drawing.unsupported`; skipped |
| VML (`w:pict`) and embedded OLE objects | Reported | `docx.vml` / `docx.embedded-object` |
| `altChunk` (embedded HTML/RTF) | Reported | `docx.altChunk` |
| Fields (`w:fldSimple`, `PAGE`, `TOC`) | Partial | Cached field text is kept; `fldChar`/`instrText` plumbing and computed values are not decoded |
| Tab stops (`w:tabs`), bookmark navigation | Partial | A `w:tab` renders as a tab character at the default interval; custom stops and bookmark jumps are not resolved |
| `.doc` (legacy binary) | Reported | `UNSUPPORTED_FORMAT` naming the legacy Word binary; an encrypted `.docx` is still `ENCRYPTED` |

## Excel workbooks (`.xlsx`, `.xlsm`)

A workbook is read from SpreadsheetML and each worksheet becomes a grid page
(hidden sheets are shown, with a diagnostic). Sheets are resolved through `xl/_rels/workbook.xml.rels`, so a
renamed or reordered package still shows the right sheets. Nothing is evaluated:
formulas display the cached result, and a formula without one is reported.

| Feature | Status | Notes |
| --- | --- | --- |
| Workbook and sheet order (`xl/workbook.xml` + relationships) | Implemented | Part names and numeric suffixes are not used for ordering |
| Shared strings, inline strings, numbers, booleans, error values | Implemented | Values are always text in the DOM, never markup |
| Number formats: general, date, date-time, percentage, custom codes | Partial | Common built-in and custom date/percent formats are recognised; an unrecognised code falls back to the raw number with a diagnostic |
| Date systems 1900 and 1904 | Implemented | The 1900 leap-year quirk is preserved |
| Formulas | Partial | The cached `<v>` is shown; formulas are never recalculated (`xlsx.formula.cache` when absent) |
| Sparse rows/columns | Implemented | Cells keep their real coordinates; empty rows are not collapsed |
| Merged ranges (`mergeCells`) | Implemented | `colspan`/`rowspan`; overlapping or hidden-content ranges are reported |
| Column widths (`cols`) | Implemented | Converted to pixels; default width otherwise |
| Hidden sheets | Implemented | Shown, with an info diagnostic naming them |
| Charts, pivot tables, conditional formatting, images, comments | Reported | `xlsx.sheet.kind` / not read; the grid is unaffected |
| Chartsheets and dialog sheets | Reported | Named in a diagnostic; not drawn |
| Macros (`.xlsm`) | Not supported | The macro part is never read or executed |
| Cell and row caps | Partial | 50 000 cells parsed, 10 000 cells rendered per page, 5 000 rows / 200 columns per sheet, 2 000 merges; each cap is reported instead of freezing |
| `.xls` (legacy binary) | Reported | `UNSUPPORTED_FORMAT` naming the legacy Excel binary |

## PDF files

PDF pages are rasterised by [pdf.js](https://mozilla.github.io/pdf.js/) onto a
canvas, one page at a time and only when visible. This is the same engine Firefox
ships; the viewer does not attempt to read PDF content streams itself.

| Feature | Status | Notes |
| --- | --- | --- |
| Page count, page sizes, first-page geometry | Implemented | Mixed page sizes are shown at the first page's size with a `pdf.page-size` diagnostic |
| Raster rendering of text, vector art, images, gradients, transparency | Implemented | By pdf.js, at the viewer's current scale × device pixel ratio |
| Zoom, fit modes, thumbnails, presentation mode, keyboard navigation | Implemented | Raster is redrawn at the new density, so pages stay sharp |
| Password-protected PDFs | Reported | `ENCRYPTED` error |
| Damaged or truncated PDFs | Reported | `MALFORMED_DOCUMENT` error |
| Linearised/encrypted-with-empty-password, object streams, JPEG 2000 | Implemented | Supported by pdf.js |
| AcroForm fields, annotations, digital signatures | Partial | Drawn as static appearance; not interactive |
| JavaScript, embedded files, portfolios | Not supported | Never executed or opened (`isEvalSupported: false`) |
| Text selection and the accessibility text layer | Not supported | One `<canvas role="img">` per page; there is no selectable text layer |
| Rendering without a canvas (SSR, some test environments) | Reported | `pdf.page` diagnostic and a labelled notice per page |
| Parsing in a Web Worker | Implemented | When `pdfWorkerSrc` is set; otherwise the main-thread fallback reports `pdf.worker.main-thread` |
| API ↔ worker version agreement | Implemented | The build (modern or `legacy`) is feature-detected, and a worker that cannot be used is reported as `pdf.engine.worker` and retried in-process instead of failing the load |
| pdf.js `legacy` build on older engines | Implemented | Reported as `pdf.engine.legacy`; polyfills included, so PDFs open in Electron-style shells and older browsers |

## Viewer chrome

| Feature | Status | Notes |
| --- | --- | --- |
| Previous/next and the slide counter | Implemented | Counter is `aria-live="polite"`; first/last are keyboard (`Home`/`End`) and handle actions rather than pill buttons |
| Floating toolbar, shown on hover | Implemented | A pill centred at the bottom of the stage; it overlays the slide instead of taking a header strip, the fit calculation reserves its height, and it is hidden (and inert) until the viewer is hovered or one of its controls has focus. Touch devices show it always |
| Thumbnails (click to navigate, lazy-mounted) | Implemented | Toggleable at runtime |
| Zoom in/out, and one dropdown of zoom stops plus fit modes | Implemented | The dropdown offers the conventional stops (10 %–500 %) and fit width/height/page/actual size; the ± buttons step a finer ladder (25 %–400 %) and `setZoom` accepts 10 %–800 % |
| Zoom and fit dropdown on click and keyboard | Implemented | Opens on click (not on hover) so a pointer crossing the pill cannot cover the slide; the percentage and its caret are one split button; `menuitemradio` semantics with the mode in effect marked; `Esc` closes and returns focus to the trigger |
| Fullscreen | Implemented | Native API first, in-page overlay fallback; `Esc` exits |
| Presentation (slideshow) mode | Implemented | Chrome and sidebar removed, black stage, click/Space/arrow/Page keys advance, `Esc` exits, controlled or uncontrolled |
| Keyboard navigation | Implemented | Arrows, PageUp/Down, Space, Home, End, `n`/`p`, `+`/`-`, `0`, `f` |
| Localisable labels | Implemented | `labels` prop, including the slide counter as a function |
| Dark chrome | Implemented | `dark` or `theme="dark"` |
| Loading, error, empty-slide and per-slide failure states | Implemented | Replaceable with `renderLoader`/`renderError`/`renderSlideError` |
| Print / PDF export | Not supported | Out of scope |
| Editing | Not supported | Out of scope |

## Images

| Feature | Status | Notes |
| --- | --- | --- |
| Detection from magic bytes | Implemented | PNG, JPEG, GIF, BMP, WebP, AVIF and text-SVG; a container signature still wins over image bytes |
| Natural size read from the header | Implemented | PNG `IHDR`, JPEG `SOFn` marker chain, GIF logical screen, BMP info header, WebP `VP8X`/`VP8 `/`VP8L`, SVG `width`/`height` or `viewBox` |
| Header without a usable size | Partial | The image is drawn in a default 1024×768 box, letterboxed, and a diagnostic says so |
| AVIF size | Partial | AVIF is detected and decoded by the browser, but its size is not read from the `ispe` box, so the page uses the default box |
| TIFF (`.tif`, `.tiff`) | Implemented | The IFD chain gives every page and its natural size; pages are decoded to RGBA and drawn on a canvas, so no browser TIFF support is needed. A page that cannot be decoded shows a labelled notice (`tiff.page`) |
| EMF / WMF | Reported | `UNSUPPORTED_FORMAT` naming the format; no browser decodes them, so they are not guessed at |
| Animation (animated GIF, animated WebP) | Implemented | The browser animates the `<img>`; the page is its first frame's size |
| EXIF orientation | Implemented | Handled by the browser (`image-orientation: from-image`), never by the viewer |
| Metadata (camera, GPS, colour profile) | Not supported | Not read, and therefore not exposed to the page |

## Delimited text (`.csv`, `.tsv`)

| Feature | Status | Notes |
| --- | --- | --- |
| Delimiter detection | Implemented | Comma, semicolon, tab or pipe; chosen by the candidate with a consistent field count across the sample, counting only delimiters outside quotes. Overridable with `contentHint` |
| Quoted fields, `""` escapes, embedded newlines | Implemented | Character-by-character state machine, not `split()` |
| CRLF / LF / CR records, missing final newline | Implemented | All four endings |
| Values kept as text | Implemented | No number, date or boolean coercion; `0123` stays `0123` |
| Rows that do not match the header | Reported | Counted in a diagnostic; values are shown as they are, never padded or shifted |
| Wide tables | Partial | A table wider than a portrait page gets a landscape page; wider than that and the columns are scaled down and clipped, with a diagnostic |
| Header row | Implemented | Repeated at the top of every page |
| Row limit | Partial | 20 000 rows kept (configurable in code, reported as a diagnostic when exceeded) |
| Encoding | Implemented | UTF-8, UTF-16 (with or without a BOM), otherwise Windows-1252 with a warning |
| Type inference, formulas, multi-line cells beyond quoting | Not supported | Out of scope: this is a viewer, not a spreadsheet |

## Text files (`.txt`, `.log`, `.json`, `.html`, `.htm`)

| Feature | Status | Notes |
| --- | --- | --- |
| Pagination | Implemented | Monospace, 12px on a 16px baseline, US Letter, wrapped to the page's column count, 1000 pages maximum |
| Line endings and tabs | Implemented | CRLF/CR/LF; tabs expanded to 4 spaces before measuring |
| Control characters | Implemented | Stripped before rendering; the text cannot smuggle terminal escapes into the DOM |
| Encoding detection | Implemented | UTF-8 (with or without a BOM), UTF-16 (with a BOM, or from its NUL pattern), otherwise Windows-1252 with a warning |
| JSON | Implemented | Pretty-printed with two-space indent up to 2 MiB; larger files are shown as they are, with a diagnostic |
| HTML | Implemented | Shown as **source text** — never parsed, never framed, never executed, and never given an origin, so nothing in it can run or load. There is no rendering option |
| Markdown | Partial | Currently shown as source text, with a diagnostic saying so |
| Syntax highlighting | Not supported | Out of scope for now |
| Text selection and search | Implemented | Real text in the DOM, so the browser's own find and select work |
| Whitespace-only or binary bytes | Reported | Refused as not a document (`NOT_A_PRESENTATION`), rather than shown as a blank page |

## Test coverage behind these claims

The matrix above is backed by the automated suite (`npm test` and
`npm run test:browser` in real Chrome) plus the browser demo. The browser suite
opens every supported type — deck, document, workbook, PDF, TIFF, raster image,
table and text — through the demo's own file picker.
Coverage includes: format detection from bytes, slide order from relationships,
theme colour resolution and modifiers, table merges, chart series and
doughnut hole sizes, group transforms, custom geometry, bullets from master styles,
missing relationships, empty slides, unsupported graphic frames, image decode
failures, hyperlink sanitisation, keyboard navigation, zoom/fit/fullscreen,
thumbnails, Strict Mode double rendering, unmount cleanup, archive limits and
ZIP-bomb rejection.

For `.docx`: styles and numbering resolution, tables with merges, inline pictures,
page breaks, keep-with-next, unit conversion, the estimated-layout fallback and the
missing-main-part error. For PDF: detection, page count and geometry from a real
PDF, the metadata title, a damaged file and an encrypted file, plus a browser test
that rasterises an actual page onto a canvas with a real worker.

For images: every supported header is measured against a fixture (including a
top-down BMP and a truncated PNG), an undecodable format is refused by name, and a
real browser test uploads a PNG and checks the page box matches it. For delimited
text: delimiter detection (comma, semicolon, tab, quoted commas, inconsistent
rows), CRLF/CR/LF endings, escaped quotes, embedded newlines, ragged rows, the row
limit, and header repetition across pages. For text files: BOM and encoding cases,
whitespace-only and binary rejection, JSON pretty-printing, wrapping, page overflow
and the page ceiling. For HTML: it is read as text with a diagnostic saying so, and
a browser test uploads a hostile page — one that scripts the host document, calls
`fetch` and loads a remote image — and asserts against the host page and the
network that none of it happened: no frame anywhere in the stage, no global
assigned, no `pageerror`, and no response at all for the beacon.

It does **not** include a corpus of real-world Office documents. Compatibility
claims therefore describe the constructs listed above, not "any .pptx, .docx,
.xlsx, .pdf or .tiff file".
