#!/usr/bin/env node
/**
 * Package verification.
 *
 * Proves that the *published artefact* — not the source tree — works:
 *
 *  1. `npm pack` the package and inspect the file list (no sources, tests or
 *     build config leaked, CSS and declarations present);
 *  2. install the tarball into a separate React 19 consumer and render a real
 *     `.pptx` from it in a DOM, asserting the slide text reaches the page;
 *  3. install it into a separate React 17 consumer and render through the CJS
 *     build to confirm the declared minimum React version works.
 *
 * Usage: node scripts/verify-package.mjs
 */

import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const workDir = join(root, ".tmp-package-verify");

function run(command, args, options = {}) {
  return execFileSync(command, args, {
    stdio: options.capture ? "pipe" : "inherit",
    encoding: "utf8",
    cwd: options.cwd ?? root,
    env: { ...process.env, npm_config_loglevel: "error" },
  });
}

function step(message) {
  process.stdout.write(`\n\u001b[1m▸ ${message}\u001b[0m\n`);
}

function fail(message) {
  process.stderr.write(`\n\u001b[31m✗ ${message}\u001b[0m\n`);
  process.exitCode = 1;
  throw new Error(message);
}

function assert(condition, message) {
  if (!condition) fail(message);
  process.stdout.write(`  ✓ ${message}\n`);
}

/* ---------------------------------------------------------------- fixtures */

/** A minimal but genuine .pptx, written by hand so no binary fixture is needed. */
const MINIMAL_PPTX_SCRIPT = `
import { zipSync, strToU8 } from "fflate";

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

export function minimalPptx() {
  const files = {
    "[Content_Types].xml": strToU8(\`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>\`),
    "_rels/.rels": strToU8(\`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>\`),
    "ppt/presentation.xml": strToU8(\`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation \${NS}>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
</p:presentation>\`),
    "ppt/_rels/presentation.xml.rels": strToU8(\`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>\`),
    "ppt/slides/slide1.xml": strToU8(\`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld \${NS}>
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:rPr lang="en-US" sz="4000" b="1"/><a:t>Packaged viewer works</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>\`),
  };
  return zipSync(files, { level: 6 });
}
`;

const ESM_CONSUMER = `
import assert from "node:assert/strict";
import { Window } from "happy-dom";
import { minimalPptx } from "./fixture.mjs";

const window = new Window({ url: "https://consumer.example/" });
const globals = [
  "window", "document", "navigator", "HTMLElement", "Element", "Node", "Text", "Comment",
  "Blob", "URL", "Event", "CustomEvent", "KeyboardEvent", "MouseEvent", "MutationObserver",
  "requestAnimationFrame", "cancelAnimationFrame", "getComputedStyle", "IntersectionObserver",
];
for (const name of globals) {
  if (window[name] === undefined) continue;
  try {
    Object.defineProperty(globalThis, name, {
      configurable: true,
      writable: true,
      value: window[name],
    });
  } catch {
    // Some Node globals (for example navigator) are getter-only; skip them.
  }
}
globalThis.IS_REACT_ACT_ENVIRONMENT = true;

// A usable container size, so the viewer is not stuck in the "no size" state.
window.HTMLElement.prototype.getBoundingClientRect = function () {
  return { x: 0, y: 0, top: 0, left: 0, right: 1200, bottom: 800, width: 1200, height: 800, toJSON: () => ({}) };
};
globalThis.ResizeObserver = class {
  observe() {}
  unobserve() {}
  disconnect() {}
};

const React = (await import("react")).default;
const { createRoot } = await import("react-dom/client");
const { act } = await import("react");
const {
  DocumentViewer,
  DocumentViewerError,
  isDocumentViewerError,
  supportedPresetNames,
  DEFAULT_ARCHIVE_LIMITS,
  PPTViewer,
  PPTViewerError,
  isPPTViewerError,
  SUPPORTED_FORMATS,
  SUPPORTED_EXTENSIONS,
  ACCEPT_ATTRIBUTE,
  missingDescriptors,
} = await import("react-anydoc-viewer");

assert.ok(Array.isArray(supportedPresetNames()) && supportedPresetNames().length > 40, "preset list exported");
assert.ok(DEFAULT_ARCHIVE_LIMITS.maxEntries > 0, "archive limits exported");
assert.ok(DocumentViewer, "DocumentViewer is exported");
assert.match(String(DocumentViewer.$$typeof), /react\.forward_ref/, "DocumentViewer is a React component type");
// The names from before the rename must survive the published build, as aliases.
assert.equal(PPTViewer, DocumentViewer, "PPTViewer is an alias of the same component");
assert.equal(PPTViewerError, DocumentViewerError, "PPTViewerError is the same class");
assert.equal(isPPTViewerError, isDocumentViewerError, "isPPTViewerError is the same guard");
assert.deepEqual(
  SUPPORTED_FORMATS,
  ["pptx", "docx", "xlsx", "pdf", "image", "csv", "text"],
  "the advertised formats are the implemented ones",
);
// The format table is the single source of truth, so a format cannot be
// advertised without an entry (extensions, MIME types) next to it.
assert.deepEqual(missingDescriptors(), [], "every advertised format has a descriptor");
assert.ok(
  SUPPORTED_EXTENSIONS.includes(".pptx") &&
    SUPPORTED_EXTENSIONS.includes(".xlsx") &&
    SUPPORTED_EXTENSIONS.includes(".tiff") &&
    SUPPORTED_EXTENSIONS.includes(".png"),
  "the extension list is exported and dot-prefixed",
);
assert.ok(
  ACCEPT_ATTRIBUTE.includes(".pptx") && ACCEPT_ATTRIBUTE.includes("application/pdf"),
  "the accept attribute is ready to paste into an <input type=file>",
);

const errors = [];
const host = window.document.createElement("div");
window.document.body.appendChild(host);
const root = createRoot(host);

await act(async () => {
  root.render(
    React.createElement(DocumentViewer, {
      source: minimalPptx(),
      showThumbnails: true,
      onError: (error) => errors.push(error),
    }),
  );
});
// Let the async load/parse settle.
for (let i = 0; i < 20 && !host.querySelector(".radv-slide"); i += 1) {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 10));
  });
}

assert.equal(errors.length, 0, "no load errors");
const rootElement = host.querySelector("[data-testid='rv-root']");
assert.ok(rootElement, "viewer root rendered");
assert.equal(rootElement.getAttribute("data-status"), "ready", "presentation reached the ready state");
assert.ok(host.querySelector(".radv-slide"), "a slide element rendered");
assert.match(host.textContent, /Packaged viewer works/, "slide text reached the DOM");
assert.match(host.textContent, /1 \\/ 1/, "slide counter rendered");
// Toolbar chrome gained its accessible name from the package's CSS-independent markup.
assert.ok(host.querySelector("[aria-label='Next slide']"), "accessible navigation control rendered");
// Thumbnails from the built package.
assert.ok(host.querySelector(".radv-thumb-item"), "thumbnail sidebar rendered");

// The zoom/fit dropdown is the newest part of the chrome, so the published bundle
// is made to prove it is there and that it opens on a click rather than on hover.
const zoomTrigger = host.querySelector("[aria-label='Zoom and fit']");
assert.ok(zoomTrigger, "zoom and fit dropdown rendered");
assert.equal(zoomTrigger.getAttribute("aria-expanded"), "false", "dropdown starts closed");
assert.ok(!host.querySelector(".radv-menu"), "no menu is rendered before the click");
await act(async () => {
  zoomTrigger.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
});
assert.equal(zoomTrigger.getAttribute("aria-expanded"), "true", "clicking the percentage opens the menu");
const menuItems = host.querySelectorAll("[role='menuitemradio']");
assert.equal(menuItems.length, 15, "the menu lists 11 zoom stops and 4 fit modes");
assert.match(host.textContent, /Fit to height/, "the new fit mode is offered by the published build");
await act(async () => {
  zoomTrigger.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }));
});
assert.equal(zoomTrigger.getAttribute("aria-expanded"), "false", "a second click closes the menu");

// Error model is exported and usable.
const err = new DocumentViewerError("NOT_A_PRESENTATION", "nope");
assert.ok(isDocumentViewerError(err), "error helpers exported");
assert.equal(err.code, "NOT_A_PRESENTATION", "error code preserved");

// Render a second, invalid source to exercise the error path in the built artefact.
const brokenHost = window.document.createElement("div");
window.document.body.appendChild(brokenHost);
const brokenRoot = createRoot(brokenHost);
const seen = [];
await act(async () => {
  brokenRoot.render(React.createElement(DocumentViewer, { source: new Uint8Array([1, 2, 3, 4]), onError: (e) => seen.push(e) }));
});
for (let i = 0; i < 20 && !brokenHost.querySelector("[data-testid='rv-error']"); i += 1) {
  await act(async () => {
    await new Promise((resolve) => setTimeout(resolve, 10));
  });
}
assert.equal(seen[0]?.code, "NOT_A_PRESENTATION", "invalid input produced a typed error");

console.log("\\n  ESM consumer (React 19): OK");
process.exit(0);
`;

const CJS17_CONSUMER = `
const assert = require("node:assert/strict");
const React = require("react");
const { renderToString } = require("react-dom/server");
const { DocumentViewer, PPTViewer } = require("react-anydoc-viewer");

assert.ok(DocumentViewer, "CJS build exports DocumentViewer");
assert.match(String(DocumentViewer.$$typeof), /react\.forward_ref/, "CJS export is a React component type");
assert.equal(PPTViewer, DocumentViewer, "CJS build keeps the deprecated alias");

const html = renderToString(
  React.createElement(DocumentViewer, {
    source: new Uint8Array([0, 1, 2, 3]),
    showThumbnails: false,
  }),
);
assert.match(html, /radv-root/, "CJS build renders the viewer root under React 17");
assert.match(html, /radv-toolbar/, "toolbar is server-rendered");
assert.match(html, /Next slide/, "accessible controls are present in SSR output");

console.log("\\n  CJS consumer (React 17, SSR): OK");
process.exit(0);
`;

/* ------------------------------------------------------------------- main */

/** `npm pack --json` can be preceded by notices, so the JSON is located first. */
function parsePackOutput(output) {
  const start = output.indexOf("[");
  const end = output.lastIndexOf("]");
  if (start === -1 || end === -1) fail(`Could not parse npm pack output:\n${output}`);
  const parsed = JSON.parse(output.slice(start, end + 1));
  const entry = Array.isArray(parsed) ? parsed[0] : parsed;
  if (!entry?.filename) fail("npm pack did not report a filename");
  return entry;
}

function writeConsumer(name, { reactVersion, scripts, packageType }) {
  const dir = join(workDir, name);
  rmSync(dir, { recursive: true, force: true });
  mkdirSync(join(dir, "src"), { recursive: true });
  writeFileSync(
    join(dir, "package.json"),
    JSON.stringify(
      {
        name: `consumer-${name}`,
        private: true,
        version: "0.0.0",
        type: packageType,
        scripts,
      },
      null,
      2,
    ),
  );
  writeFileSync(join(dir, "react-version.txt"), reactVersion);
  return dir;
}

function install(dir, tarball, dependencies) {
  run("npm", ["install", "--no-audit", "--no-fund", "--no-package-lock", tarball, ...dependencies], { cwd: dir });
}

async function main() {
  rmSync(workDir, { recursive: true, force: true });
  mkdirSync(workDir, { recursive: true });

  step("Packing the library");
  const packJson = run("npm", ["pack", "--json", "--pack-destination", workDir], { capture: true });
  const entry = parsePackOutput(packJson);
  const tarball = join(workDir, entry.filename);
  assert(existsSync(tarball), `tarball created: ${entry.filename}`);

  const files = entry.files.map((file) => file.path);
  assert(files.some((file) => file === "dist/index.js"), "ESM build published");
  assert(files.some((file) => file === "dist/index.cjs"), "CJS build published");
  assert(files.some((file) => file === "dist/index.d.ts"), "type declarations published");
  assert(files.some((file) => file === "dist/styles.css"), "stylesheet published");
  assert(files.some((file) => file.toLowerCase() === "license"), "LICENSE published");
  assert(files.some((file) => file.startsWith("README")), "README published");
  assert(!files.some((file) => file.startsWith("src/")), "no source files leaked into the tarball");
  assert(!files.some((file) => file.startsWith("tests/")), "no test files leaked into the tarball");
  assert(!files.some((file) => file.startsWith("examples/")), "no example files leaked into the tarball");
  assert(
    !files.some((file) => file.includes(".tmp-package-verify")),
    "no verification scratch files leaked",
  );
  process.stdout.write(`  published files: ${files.length}\n`);

  step("Installing into a React 19 ESM consumer");
  const esmDir = writeConsumer("esm", {
    reactVersion: "19",
    packageType: "module",
    scripts: { start: "node --experimental-vm-modules ./src/main.mjs" },
  });
  writeFileSync(join(esmDir, "src/fixture.mjs"), MINIMAL_PPTX_SCRIPT);
  writeFileSync(join(esmDir, "src/main.mjs"), ESM_CONSUMER);
  install(esmDir, tarball, ["react@^19", "react-dom@^19", "happy-dom@^21", "fflate@^0.8"]);
  run("node", ["./src/main.mjs"], { cwd: esmDir });

  step("Installing into a React 17 CJS consumer");
  const cjsDir = writeConsumer("cjs17", {
    reactVersion: "17",
    packageType: "commonjs",
    scripts: { start: "node ./src/main.cjs" },
  });
  writeFileSync(join(cjsDir, "src/main.cjs"), CJS17_CONSUMER);
  install(cjsDir, tarball, ["react@^17.0.2", "react-dom@^17.0.2"]);
  run("node", ["./src/main.cjs"], { cwd: cjsDir });

  step("Checking the published stylesheet and subpath export");
  const installedCss = join(esmDir, "node_modules/react-anydoc-viewer/dist/styles.css");
  assert(existsSync(installedCss), "subpath 'react-anydoc-viewer/styles.css' resolves to a real file");
  const css = readFileSync(installedCss, "utf8");
  assert(css.includes(".radv-root"), "stylesheet contains the viewer root styles");
  assert(css.includes(".radv-theme-dark"), "stylesheet ships the dark theme");

  step("Type-checking a consumer against the published declarations");
  const typecheckDir = join(workDir, "types");
  rmSync(typecheckDir, { recursive: true, force: true });
  mkdirSync(join(typecheckDir, "src"), { recursive: true });
  writeFileSync(
    join(typecheckDir, "package.json"),
    JSON.stringify({ name: "consumer-types", private: true, version: "0.0.0", type: "module" }, null, 2),
  );
  writeFileSync(
    join(typecheckDir, "tsconfig.json"),
    JSON.stringify(
      {
        compilerOptions: {
          target: "ES2020",
          module: "ESNext",
          moduleResolution: "bundler",
          jsx: "react-jsx",
          strict: true,
          noEmit: true,
          skipLibCheck: false,
          lib: ["ES2021", "DOM"],
        },
        include: ["src"],
      },
      null,
      2,
    ),
  );
  writeFileSync(
    join(typecheckDir, "src/main.tsx"),
    `import { createRef } from "react";
import { DocumentViewer, DocumentViewerError, isDocumentViewerError } from "react-anydoc-viewer";
import type { DocumentViewerHandle, DocumentViewerProps, DocumentInfo, DocumentDiagnostic } from "react-anydoc-viewer";
// The deprecated names must still be importable *and* typed in the published
// declarations, or a consumer's upgrade would fail at compile time.
import { PPTViewer, PPTViewerError, isPPTViewerError, detectPdfEngine } from "react-anydoc-viewer";
import type { PPTViewerHandle, PPTViewerProps, PptxSource, PresentationInfo } from "react-anydoc-viewer";
import type { PdfEngine } from "react-anydoc-viewer";
import "react-anydoc-viewer/styles.css";

const ref = createRef<DocumentViewerHandle>();

export const props: DocumentViewerProps = {
  file: "/deck.pptx",
  showControls: true,
  showThumbnails: true,
  fit: "width",
  onLoad: (info: DocumentInfo) => void info.slideCount,
  onError: (error) => {
    const typed: DocumentViewerError = error;
    if (isDocumentViewerError(typed)) void typed.code;
  },
  onSlideChange: (index: number) => void index,
  onDiagnostics: (diagnostics: DocumentDiagnostic[]) => void diagnostics.length,
};

export const element = <DocumentViewer {...props} ref={ref} />;

// Deprecated names: the same component, the same ref and the same types, so an
// un-migrated file still compiles and renders.
const legacyRef = createRef<PPTViewerHandle>();
export const legacyProps: PPTViewerProps = { file: "/report.docx" as PptxSource };
export const legacyElement = <PPTViewer {...legacyProps} ref={legacyRef} />;
export const legacyInfo = (info: PresentationInfo): string => info.format;
export const engine: PdfEngine = detectPdfEngine();
export const legacyError: PPTViewerError = new PPTViewerError("NO_SLIDES", "empty");
export const stillGuarded = isPPTViewerError(legacyError);
`,
  );
  // React's own types must be installed explicitly. Without them TypeScript walks
  // up the directory tree and finds *this repository's* `node_modules/@types/react`,
  // so the check would pass in the repo and fail everywhere else (TS7016) — proving
  // nothing about the published declarations.
  install(typecheckDir, tarball, [
    "react@^19",
    "react-dom@^19",
    "@types/react@^19",
    "@types/react-dom@^19",
    "typescript@^5.9",
  ]);
  run("npx", ["tsc", "-p", "tsconfig.json"], { cwd: typecheckDir });
  process.stdout.write("  consumer type-check: OK\n");

  step("Verification summary");
  process.stdout.write(
    `  tarball: ${entry.filename} (${Math.round((entry.size ?? 0) / 1024)} KiB unpacked: ${entry.unpackedSize ?? "?"} bytes)\n`,
  );
  process.stdout.write("\n\u001b[32mAll package checks passed.\u001b[0m\n");
}

main().catch((error) => {
  process.stderr.write(`\n${error?.stack ?? error}\n`);
  process.exit(1);
});
