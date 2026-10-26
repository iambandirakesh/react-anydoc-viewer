# Security

`react-anydoc-viewer` renders files it did not create, usually uploaded by end
users. It treats every document it opens — `.pptx`, `.docx`, `.xlsx`, `.pdf`,
TIFF and other images, tables and text files — as hostile input, and the same limit
checks apply to all of them.

## Reporting a vulnerability

Open a private security advisory on the repository (or email the maintainer listed
in `package.json`) rather than a public issue. Please include the smallest `.pptx`
that reproduces the problem, the viewer version, and the browser. A first response
should be expected within a few days; there is no bug-bounty programme.

Supported versions: the latest release. This is a pre-1.0 library, so fixes ship as
patch or minor releases rather than backports.

## Threat model

| Threat | Mitigation |
| --- | --- |
| **ZIP bombs / resource exhaustion** | Limits are checked against the ZIP central directory *before* inflation: max entries (8000), max entry size (128 MiB), max total (512 MiB) and a compression-ratio heuristic (200:1). Exceeding any of them fails the load with `LIMIT_EXCEEDED` instead of allocating. All of them are configurable through the `limits` prop — lower them for untrusted public uploads. |
| **Path traversal / zip slip** | Entry names are validated before use: absolute paths, drive letters, `..` segments, backslashes, control characters and NUL bytes are rejected. Nothing is ever written to the filesystem, and part names are only used as keys into the archive. |
| **XML entity expansion (billion laughs) / XXE** | Any document containing a DOCTYPE or ENTITY declaration is refused outright; the XML parser is not allowed to resolve external entities. Only the five predefined entities and numeric character references are decoded, manually. |
| **Oversized XML parts** | `maxXmlBytes` (48 MiB) caps a single decoded part. |
| **Script execution from document content** | Document content is never evaluated. Text is rendered as React text nodes (escaped), so markup in a page is displayed literally. There is no `dangerouslySetInnerHTML` anywhere in the library. |
| **Script execution inside a PDF (JavaScript actions, XFA, embedded files)** | pdf.js is configured with `isEvalSupported: false`, and the viewer renders pages as raster output only: no annotation, form or JavaScript layer is activated, no embedded file is opened, and a PDF's `/OpenAction` is never run. Links inside a PDF are not clickable, because the rendered canvas carries no link annotations. |
| **Exhausting memory through a raster page** | The canvas backing store is bounded: density is clamped to 3× CSS pixel density and a page to 8 megapixels, so a PDF declaring an enormous `/MediaBox` cannot allocate an unbounded bitmap. Pages are rasterised one at a time and cancelled when they leave. |
| **Malicious hyperlinks** | Only `http`, `https`, `mailto` and `tel` become `href`s, and only when rendering interactively. `javascript:`, `data:`, `vbscript:`, `file:` and anything unknown are stripped and reported as `hyperlink.blocked`; the text remains visible but is not a link. |
| **SVG with embedded scripts** | SVG media is always served through a blob/`data:` URL in an `<img>` element, where scripts do not execute and the SVG cannot reach the host document. |
| **Cross-site data leakage** | The library makes no network requests of its own. Fetching happens only for a `source` you pass (with your `fetchOptions`, so you control credentials and CORS), and downloads are capped by `maxDownloadBytes`. |
| **HTML files** | Shown as source text, paginated like any other text file: never parsed, never framed, never given an origin, so nothing in it can run or load. There is no rendering option to misuse — no iframe, no `srcdoc`, no `innerHTML` — and the file's own content is displayed as characters like any other text. An HTML file is the one format with no safe way to be *executed* here, so it is not executed at all. |
| **Script from an uploaded image** | Every browser-decodable image, SVG included, is delivered as a blob/`data:` URL in an `<img>` element: scripts do not run there and the file cannot reach the host document. A TIFF is decoded by the viewer's own bounded reader and drawn as pixels on a canvas, never as markup. Vector metafiles a browser cannot decode (EMF, WMF) are refused with a named error. |
| **Runaway work from a workbook** | Worksheet XML goes through the same archive limits and the same DOCTYPE/ENTITY refusal as every other part. Cells are capped (50 000 parsed, 10 000 rendered per sheet page), as are rows, columns and merged ranges, and hitting a cap emits a diagnostic instead of exhausting memory. Formula text is never evaluated; only the cached value is displayed. |
| **Runaway work from a plain text or table file** | Text pages are capped at 1000 pages; a table keeps at most 20 000 rows and parses at most 8 MiB of characters. Both truncations are reported as diagnostics, so a host can refuse the file instead of accepting a partial view. Decoding is bounded by the archive limits and `maxDownloadBytes` for downloads. |
| **Misleading decoding of text** | A file that is not valid UTF-8 is decoded as Windows-1252 and reported: the viewer never silently substitutes characters, because a document that displays the wrong text without saying so is a security problem as well as a bug. |
| **Table content executed as a formula** | Cell values are rendered as text, never interpreted. Values that begin with `=`, `+`, `-` or `@` are displayed literally; a host that exports them back to a spreadsheet must neutralise them itself (formula injection is an export concern). |
| **Resource leaks / memory growth** | Object URLs are tracked and revoked on unmount and on source change; in-flight loads are aborted (superseded loads never render); a destroyed PDF document proxy is released with the load; media is inflated on demand. |
| **Denial of service through pathological files** | Deeply nested groups, thousands of shapes, huge tables, tens of thousands of paragraphs and pages that are never visited are bounded by the archive limits and by lazy rendering; unbounded recursion is not used when walking shape trees, and a page that throws is contained by an error boundary so one bad page cannot break the viewer. |
| **Supply chain** | Runtime dependencies are `fflate` (MIT), `fast-xml-parser` (MIT), `pdfjs-dist` (Apache-2.0, Mozilla — the engine behind Firefox's PDF viewer) and `utif` (MIT, a small TIFF decoder). Nothing is fetched from a CDN at runtime, the package ships no post-install scripts, and both `pdfjs-dist` and `utif` are imported lazily, so they are only evaluated when a PDF or a TIFF is opened. |

## What the library does *not* protect against

- **Host application mistakes.** Sanitising the surrounding page, CSP, and the
  decision to render arbitrary user uploads at all remain the application's
  responsibility.
- **A page you decide to render yourself.** This library will not execute an
  uploaded HTML file, but if your application does — in an iframe, a preview pane
  or an editor — that decision and its isolation are yours. Serve untrusted uploads
  from a separate cookie-less origin: it is a stronger boundary than any iframe
  attribute, and sandboxing does not stop the framed document's own network loads.
- **Content-level deception.** A presentation can contain misleading text, images
  or hyperlinks. The viewer renders what the file says.
- **Rendering engines.** Fonts, images and text layout are the browser's job, PDF
  rendering is pdf.js's, and TIFF pixel decoding is `utif`'s. A vulnerability in the
  browser's image or font decoders, in pdf.js or in utif is a vulnerability of that
  library; keep `pdfjs-dist` and `utif` up to date and treat a `.pdf` or `.tiff` from
  an untrusted source with the same care as any file you would open in a browser.
- **Encrypted or macro-enabled files.** Password-protected packages and VBA
  macros are rejected or ignored, not analysed: a `.pptm`/`.docm` opens as a normal
  file with its macros never executed, and an encrypted package is refused with an
  `ENCRYPTED` error. A PDF with an empty-password encryption dictionary is opened by
  pdf.js without prompting, as a desktop reader would.

## Hardening checklist for integrators

```tsx
const limits = {
  maxCompressedBytes: 16 * 1024 * 1024, // 16 MiB uploads
  maxEntries: 2000,
  maxTotalBytes: 128 * 1024 * 1024,
  maxCompressionRatio: 50,
};

<DocumentViewer
  source={userFile}
  limits={limits}
  fetchOptions={{ credentials: "omit", mode: "cors" }}
  maxDownloadBytes={16 * 1024 * 1024}
/>
```

- Lower the limits for user-uploaded content.
- Links are always sanitised; if links in untrusted documents are undesirable at
  all, strip the `<a>` elements from the rendered DOM after mount (the library never
  emits an `href` it has not validated).
- Keep the dependencies up to date; `npm audit` covers the four runtime
  dependencies, and `pdfjs-dist` and `utif` in particular parse attacker-controlled
  file formats directly.
- Serve the host page with a restrictive CSP: the library needs no `unsafe-eval`,
  no inline scripts and no external origins. When you pass `pdfWorkerSrc`, allow the
  same origin for workers (`worker-src 'self'`).
