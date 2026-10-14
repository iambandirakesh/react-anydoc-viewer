// @vitest-environment node
/**
 * The PDF adapter runs in a Node environment on purpose: it needs no DOM (pages
 * are rasterised later, in the browser) and this is where pdf.js's own Node build
 * can be exercised for real. `parsePdfDocument` imports pdf.js lazily, so these
 * tests also prove the module boundary works — nothing is loaded until a PDF is
 * actually parsed.
 */

import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import { describe, expect, it, vi } from "vitest";
import { detectFormat, parseDocumentBytes } from "../../src/parser/document";
import { detectPdfEngine, normalizePdfError, pdfEngineChoice } from "../../src/parser/pdf";
import type { PdfParseOptions } from "../../src/parser/pdf";
import { emuToPx } from "../../src/parser/units";
import { buildFakePdf, buildPdf } from "../fixtures/pdf";

const require = createRequire(import.meta.url);

/**
 * pdf.js needs a worker even in Node, where the "worker" is imported into the
 * main thread. The legacy build is the Node-compatible one.
 */
const nodeOptions: PdfParseOptions = {
  workerSrc: pathToFileURL(require.resolve("pdfjs-dist/legacy/build/pdf.worker.mjs")).href,
  loadPdfJs: () => import("pdfjs-dist/legacy/build/pdf.mjs") as never,
};

const LETTER_WIDTH_EMU = Math.round(612 * 12_700); // 8.5in at 96dpi

describe("PDF documents", () => {
  it("parses pages, page geometry and metadata", async () => {
    const parsed = await parseDocumentBytes(
      buildPdf({ pages: ["First page", "Second page"], title: "Fixture PDF", author: "Fixture author" }),
      nodeOptions,
    );
    expect(parsed.format).toBe("pdf");
    expect(parsed.model.slides).toHaveLength(2);
    expect(parsed.model.width).toBe(LETTER_WIDTH_EMU);
    expect(parsed.model.height).toBe(Math.round(792 * 12_700));
    expect(parsed.model.title).toBe("Fixture PDF");
    expect(parsed.model.author).toBe("Fixture author");
    // Pages carry no shapes: the page renderer draws them.
    expect(parsed.model.slides[0]?.shapes).toHaveLength(0);
    expect(typeof parsed.renderPage).toBe("function");
    expect(parsed.media.activeUrls).toBe(0);
    parsed.dispose();
  });

  it("reads each page's own MediaBox", async () => {
    const parsed = await parseDocumentBytes(
      buildPdf({ pages: [{ text: "A4", width: 595.276, height: 841.89 }] }),
      nodeOptions,
    );
    expect(emuToPx(parsed.model.width)).toBeCloseTo(595.276 * (96 / 72), 1);
    expect(emuToPx(parsed.model.height)).toBeCloseTo(841.89 * (96 / 72), 1);
    parsed.dispose();
  });

  it("reports pages whose size differs from the first page", async () => {
    const parsed = await parseDocumentBytes(
      buildPdf({
        pages: [
          { text: "letter" },
          { text: "tabloid", width: 792, height: 1224 },
        ],
      }),
      nodeOptions,
    );
    expect(parsed.model.slides).toHaveLength(2);
    expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain("pdf.page-size");
    parsed.dispose();
  });

  it("refuses a damaged PDF with a normalised error", async () => {
    await expect(parseDocumentBytes(buildFakePdf(), nodeOptions)).rejects.toMatchObject({
      code: "MALFORMED_DOCUMENT",
    });
    await expect(parseDocumentBytes(buildPdf({ broken: true }), nodeOptions)).rejects.toMatchObject({
      code: "MALFORMED_DOCUMENT",
    });
  });

  it("reports a missing pdf.js module instead of failing obscurely", async () => {
    const loader = () => {
      throw new Error("Cannot find module 'pdfjs-dist'");
    };
    await expect(
      parseDocumentBytes(buildPdf({ pages: ["x"] }), { ...nodeOptions, loadPdfJs: loader as never }),
    ).rejects.toMatchObject({ code: "UNKNOWN" });
  });

  it("detects PDFs without parsing them", () => {
    expect(detectFormat(buildPdf({ pages: ["x"] }))).toBe("pdf");
  });
});

/**
 * The primitives the modern pdf.js build needs, as this environment may or may
 * not provide them (Node only gained `Uint8Array.prototype.toHex` recently). The
 * returned function puts the realm back exactly as it was.
 */
function installModernApis(): () => void {
  const proto = Uint8Array.prototype as unknown as Record<string, unknown>;
  const globals = globalThis as unknown as Record<string, unknown>;
  const promise = Promise as unknown as Record<string, unknown>;
  const url = URL as unknown as Record<string, unknown>;
  const saved = {
    toHex: proto.toHex,
    float16: globals.Float16Array,
    promiseTry: promise.try,
    urlParse: url.parse,
  };
  proto.toHex = function toHex(): string {
    return "";
  };
  globals.Float16Array = function Float16Array(): void {};
  promise.try = function tryFn(fn: () => unknown): unknown {
    return Promise.resolve().then(fn);
  };
  url.parse = function parse(): unknown {
    return null;
  };
  return () => {
    for (const [key, value] of Object.entries(saved)) {
      if (value === undefined) continue;
      if (key === "toHex") proto.toHex = value;
      if (key === "float16") globals.Float16Array = value;
      if (key === "promiseTry") promise.try = value;
      if (key === "urlParse") url.parse = value;
    }
    vi.unstubAllGlobals();
  };
}

describe("pdf.js build selection", () => {
  it("uses the modern build when every modern-api feature is present", async () => {
    const restore = installModernApis();
    vi.resetModules();
    try {
      const fresh = await import("../../src/parser/pdf");
      expect(fresh.missingModernPdfFeatures()).toEqual([]);
      expect(fresh.detectPdfEngine()).toBe("modern");
    } finally {
      restore();
      vi.resetModules();
    }
  });

  it("falls back to the legacy build when a modern-api feature is missing", async () => {
    const restore = installModernApis();
    // The build that actually failed in an older Chromium: its worker calls
    // `Uint8Array.prototype.toHex`, which only recent engines implement.
    (Uint8Array.prototype as unknown as Record<string, unknown>).toHex = undefined;
    vi.resetModules();
    try {
      const fresh = await import("../../src/parser/pdf");
      expect(fresh.missingModernPdfFeatures()).toContain("Uint8Array.prototype.toHex");
      expect(fresh.detectPdfEngine()).toBe("legacy");
    } finally {
      restore();
      vi.resetModules();
    }
  });

  it("freezes the answer, because the legacy build polyfills what it detects", () => {
    const first = pdfEngineChoice();
    const restore = installModernApis();
    try {
      // A legacy load has just polyfilled the missing primitives; the choice must
      // not silently flip to a build the page cannot run.
      expect(pdfEngineChoice()).toBe(first);
      expect(detectPdfEngine()).toBe(first.engine);
      expect(first.engine).toBe(first.missing.length === 0 ? "modern" : "legacy");
    } finally {
      restore();
    }
  });

  it("parses with the legacy build when it is selected", async () => {
    const parsed = await parseDocumentBytes(buildPdf({ pages: ["legacy page"] }), {
      loadPdfJs: nodeOptions.loadPdfJs,
      pdfEngine: "legacy",
    });
    expect(parsed.model.slides).toHaveLength(1);
    expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain(
      "pdf.engine.legacy",
    );
    parsed.dispose();
  });

  it("renders on the main thread when the configured worker cannot be used", async () => {
    const attempts: Array<{ workerSrc: string; isCopy: boolean }> = [];
    const bytes = buildPdf({ pages: ["retry"] });
    const fake = {
      GlobalWorkerOptions: { workerSrc: "" },
      getDocument(params: Record<string, unknown>) {
        const first = attempts.length === 0;
        attempts.push({
          workerSrc: fake.GlobalWorkerOptions.workerSrc,
          // pdf.js detaches the buffer it is handed, so the viewer must pass a copy.
          isCopy: (params.data as Uint8Array).buffer !== bytes.buffer,
        });
        return {
          promise: first
            ? Promise.reject(
                Object.assign(
                  new Error('The API version "5.7.284" does not match the Worker version "5.2.0".'),
                  { name: "UnknownErrorException" },
                ),
              )
            : Promise.resolve({
                numPages: 1,
                getPage: () =>
                  Promise.resolve({ getViewport: () => ({ width: 612, height: 792 }) }),
                getMetadata: () => Promise.resolve({ info: { Title: "Retry" } }),
                destroy: () => Promise.resolve(),
              }),
        };
      },
    };

    const parsed = await parseDocumentBytes(bytes, {
      loadPdfJs: () => fake as never,
      // The viewer's public prop is `pdfWorkerSrc`; the adapter's own option is
      // `workerSrc`, and the document layer is what translates between them.
      pdfWorkerSrc: "https://cdn.example/pdf.worker.min.mjs",
    });
    try {
      expect(attempts).toHaveLength(2);
      expect(attempts[0]?.workerSrc).toBe("https://cdn.example/pdf.worker.min.mjs");
      expect(attempts[1]?.workerSrc).toBe("");
      expect(attempts[0]?.isCopy).toBe(true);
      expect(parsed.model.title).toBe("Retry");
      expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain(
        "pdf.engine.worker",
      );
    } finally {
      parsed.dispose();
      delete (globalThis as { pdfjsWorker?: unknown }).pdfjsWorker;
    }
  });

  it("does not retry a file the engine has rejected", async () => {
    const attempts: number[] = [];
    const fake = {
      GlobalWorkerOptions: { workerSrc: "" },
      getDocument() {
        attempts.push(1);
        return {
          promise: Promise.reject(Object.assign(new Error("Invalid PDF structure"), { name: "InvalidPDFException" })),
        };
      },
    };
    await expect(
      parseDocumentBytes(buildFakePdf(), {
        loadPdfJs: () => fake as never,
        pdfWorkerSrc: "https://cdn.example/pdf.worker.min.mjs",
      }),
    ).rejects.toMatchObject({ code: "MALFORMED_DOCUMENT" });
    expect(attempts).toHaveLength(1);
  });
});

describe("PDF error mapping", () => {
  it("maps password protection and format damage onto public codes", () => {
    const password = Object.assign(new Error("No password given"), { name: "PasswordException" });
    expect(normalizePdfError(password)).toMatchObject({ code: "ENCRYPTED" });
    const invalid = Object.assign(new Error("Invalid PDF structure"), { name: "InvalidPDFException" });
    expect(normalizePdfError(invalid)).toMatchObject({ code: "MALFORMED_DOCUMENT" });
    expect(normalizePdfError(new Error("boom"))).toMatchObject({ code: "UNKNOWN" });
    expect(normalizePdfError(Object.assign(new Error("x"), { name: "AbortError" })).code).toBe("ABORTED");
  });
});
