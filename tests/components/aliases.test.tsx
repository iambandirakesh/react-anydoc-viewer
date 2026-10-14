/**
 * The compatibility contract behind the rename.
 *
 * The package was `react-pptx-viewer` while it only opened presentations. It now
 * renders `.docx` and `.pdf` too, so the component, the types and the errors carry
 * format-neutral names — and every old name is an *alias*, not a copy. These tests
 * pin that down, because a re-export that quietly became a second component or a
 * second class would still type-check while breaking `instanceof`, refs and DOM
 * identity for everyone who has not migrated yet.
 */

import { createRef } from "react";
import { afterEach, describe, expect, it } from "vitest";
import { cleanup, render, waitFor } from "@testing-library/react";
import { buildPptx, paragraph, run, sp } from "../fixtures/pptx";
import * as api from "../../src/index";
import { DocumentViewer, PPTViewer } from "../../src/components/DocumentViewer";
import {
  DocumentViewerError,
  PPTViewerError,
  isDocumentViewerError,
  isPPTViewerError,
} from "../../src/errors";
import type {
  DocumentViewerHandle,
  DocumentViewerProps,
  PresentationInfo,
  PPTViewerHandle,
  PPTViewerProps,
  PptxDiagnostic,
  PptxSource,
} from "../../src/types/public";

afterEach(cleanup);

const oneSlide = () =>
  buildPptx({
    slides: [
      {
        part: "ppt/slides/slide1.xml",
        shapes: sp({ ph: "title", text: paragraph({ runs: run({ text: "Renamed, not replaced", size: 2400 }) }) }),
      },
    ],
  });

describe("deprecated names", () => {
  it("exposes the same component under both names", () => {
    expect(PPTViewer).toBe(DocumentViewer);
    expect(api.PPTViewer).toBe(api.DocumentViewer);
    expect(api.default).toBe(api.DocumentViewer);
  });

  it("exposes the same error class, guard and code type", () => {
    expect(PPTViewerError).toBe(DocumentViewerError);
    expect(isPPTViewerError).toBe(isDocumentViewerError);

    const error = new PPTViewerError("NOT_A_PRESENTATION", "nope");
    expect(error).toBeInstanceOf(DocumentViewerError);
    expect(isDocumentViewerError(error)).toBe(true);
    expect(error.code).toBe("NOT_A_PRESENTATION");
    expect(error.name).toBe("DocumentViewerError");
  });

  it("renders and drives the viewer through the old name", async () => {
    const ref = createRef<PPTViewerHandle>();
    const { container } = render(<PPTViewer ref={ref} source={oneSlide()} showThumbnails />);

    await waitFor(() => {
      expect(container.querySelector("[data-testid='rv-root']")?.getAttribute("data-status")).toBe("ready");
    });
    expect(container.querySelector(".radv-canvas")?.textContent).toContain("Renamed, not replaced");
    expect(ref.current?.getSlideCount()).toBe(1);
    // The ref type and the props type are aliases of the format-neutral ones.
    const handle: DocumentViewerHandle | null = ref.current;
    expect(handle?.getCurrentSlide()).toBe(0);
  });

  it("keeps the deprecated type aliases assignable to the new types", () => {
    // Compile-time assertions: these only need to type-check, so nothing runs.
    const props: DocumentViewerProps = { showControls: true } satisfies PPTViewerProps;
    const source: PptxSource = new Uint8Array([1, 2, 3]);
    const info = (value: PresentationInfo): string => value.format;
    const diagnostics: PptxDiagnostic[] = [];
    expect([props.showControls, source.byteLength, typeof info, diagnostics.length]).toEqual([
      true,
      3,
      "function",
      0,
    ]);
  });

  it("exposes the format-neutral names on the package entry point", () => {
    expect(typeof api.DocumentViewer).toBe("object");
    expect(api.SUPPORTED_FORMATS).toEqual(["pptx", "docx", "pdf"]);
    expect(typeof api.detectPdfEngine()).toBe("string");
    expect(api.DEFAULT_ARCHIVE_LIMITS.maxEntries).toBeGreaterThan(0);
  });

  it("prefixes every rendered class with radv-", async () => {
    const { container } = render(<DocumentViewer source={oneSlide()} showThumbnails />);
    await waitFor(() => {
      expect(container.querySelector("[data-testid='rv-root']")).toBeTruthy();
    });
    const classes = [...container.querySelectorAll("*")]
      .flatMap((element) => [...element.classList])
      .filter((name) => !name.startsWith("radv-"));
    expect(classes).toEqual([]);
    expect(container.querySelector(".radv-root")).toBeTruthy();
  });
});
