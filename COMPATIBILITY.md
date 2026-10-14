# Compatibility matrix

What `react-anydoc-viewer` renders today, what it renders approximately, and what it
deliberately reports as unsupported — for all three formats: PowerPoint (`.pptx`),
Word (`.docx`) and PDF.

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
| Package kind from part names (`ppt/presentation.xml`, `word/document.xml`) | Implemented | A renamed or extension-less package still opens correctly |
| Other OPC packages (`.xlsx`, OpenDocument) | Reported | `UNSUPPORTED_FORMAT` naming what the package actually is |
| Slide order via `p:sldIdLst` + `r:id` | Implemented | File names and numeric suffixes are never used for ordering |
| Slide size (`p:sldSz` cx/cy) and aspect ratio | Implemented | Slide is never stretched; `fit` scales it uniformly |
| Slide master, layout and theme resolution | Implemented | Layout → master → theme chain per slide |
| Multiple masters/layouts | Partial | All layouts of the first master are read; a deck with several masters uses the first |
| Core properties (title) | Implemented | Exposed on `onLoad` |
| Notes slides, comments, custom XML, embedded workbooks | Not supported | Never inflated (they are not in the useful-parts filter) |
| Encrypted / password-protected packages | Reported | `ENCRYPTED` error (detected from the OLE/CFB signature) |
| Legacy binary `.ppt` / `.doc` | Reported | `ENCRYPTED` error, because the OLE container is indistinguishable from an encrypted OOXML file |
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
| EMF/WMF/TIFF | Reported | Browsers cannot decode them; labelled placeholder |
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
| `.doc` (legacy binary) | Reported | `ENCRYPTED` error from the OLE signature |

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
| First/previous/next/last, slide counter | Implemented | Counter is `aria-live="polite"` |
| Thumbnails (click to navigate, lazy-mounted) | Implemented | Toggleable at runtime |
| Zoom in/out, fit page, fit width, actual size | Implemented | Zoom ladder 25 %–400 % via buttons, 10 %–800 % via `setZoom` |
| Fullscreen | Implemented | Native API first, in-page overlay fallback; `Esc` exits |
| Presentation (slideshow) mode | Implemented | Chrome and sidebar removed, black stage, click/Space/arrow/Page keys advance, `Esc` exits, controlled or uncontrolled |
| Keyboard navigation | Implemented | Arrows, PageUp/Down, Space, Home, End, `n`/`p`, `+`/`-`, `0`, `f` |
| Localisable labels | Implemented | `labels` prop, including the slide counter as a function |
| Dark chrome | Implemented | `dark` or `theme="dark"` |
| Loading, error, empty-slide and per-slide failure states | Implemented | Replaceable with `renderLoader`/`renderError`/`renderSlideError` |
| Print / PDF export | Not supported | Out of scope |
| Editing | Not supported | Out of scope |

## Test coverage behind these claims

The matrix above is backed by the automated suite (`npm test`, 147 tests, and
`npm run test:browser`, 16 Playwright tests in real Chrome) plus the browser demo.
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

It does **not** include a corpus of real-world Office documents. Compatibility
claims therefore describe the constructs listed above, not "any .pptx, .docx or
.pdf file".
