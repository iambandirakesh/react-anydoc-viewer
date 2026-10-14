# Security

`react-anydoc-viewer` renders files it did not create, usually uploaded by end
users. It treats every document it opens — `.pptx`, `.docx` or `.pdf` — as hostile
input, and the same limits apply to all three.

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
| **Resource leaks / memory growth** | Object URLs are tracked and revoked on unmount and on source change; in-flight loads are aborted (superseded loads never render); a destroyed PDF document proxy is released with the load; media is inflated on demand. |
| **Denial of service through pathological files** | Deeply nested groups, thousands of shapes, huge tables, tens of thousands of paragraphs and pages that are never visited are bounded by the archive limits and by lazy rendering; unbounded recursion is not used when walking shape trees, and a page that throws is contained by an error boundary so one bad page cannot break the viewer. |
| **Supply chain** | Runtime dependencies are `fflate` (MIT), `fast-xml-parser` (MIT) and `pdfjs-dist` (Apache-2.0, Mozilla — the engine behind Firefox's PDF viewer). Nothing is fetched from a CDN at runtime, the package ships no post-install scripts, and `pdfjs-dist` is imported lazily, so it is only ever evaluated when a PDF is opened. |

## What the library does *not* protect against

- **Host application mistakes.** Sanitising the surrounding page, CSP, and the
  decision to render arbitrary user uploads at all remain the application's
  responsibility.
- **Content-level deception.** A presentation can contain misleading text, images
  or hyperlinks. The viewer renders what the file says.
- **Rendering engines.** Fonts, images and text layout are the browser's job, and
  PDF rendering is pdf.js's. A vulnerability in the browser's image or font
  decoders, or in pdf.js itself, is a vulnerability of that library; keep
  `pdfjs-dist` up to date and treat a `.pdf` from an untrusted source with the same
  care as any PDF you would open in a browser.
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
- Keep the dependencies up to date; `npm audit` covers the three runtime
  dependencies, and `pdfjs-dist` in particular is the one that parses an
  attacker-controlled file format directly.
- Serve the host page with a restrictive CSP: the library needs no `unsafe-eval`,
  no inline scripts and no external origins. When you pass `pdfWorkerSrc`, allow the
  same origin for workers (`worker-src 'self'`).
