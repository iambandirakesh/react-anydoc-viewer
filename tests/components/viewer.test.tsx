import { StrictMode, createRef } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { act, cleanup, fireEvent, render, screen, waitFor, within } from "@testing-library/react";
import {
  buildPptx,
  inch,
  NS,
  paragraph,
  pic,
  RED_PNG,
  run,
  sp,
} from "../fixtures/pptx";
import { buildDocx, paragraph as docxParagraph, wParagraph, wRun } from "../fixtures/docx";
import { createEmptyMediaStore } from "../../src/parser/media";
import type { PageRenderer } from "../../src/renderer/pageRenderer";
import { DocumentViewer } from "../../src/components/DocumentViewer";
import { SlideCanvas } from "../../src/components/SlideCanvas";
import { mergeLabels } from "../../src/components/labels";
import { setZeroSize } from "../setup";
import type { DocumentViewerHandle } from "../../src/types/public";

afterEach(cleanup);

const twoSlides = () =>
  buildPptx({
    slides: [
      {
        part: "ppt/slides/slide1.xml",
        shapes: sp({ ph: "title", text: paragraph({ runs: run({ text: "Slide one", size: 3200 }) }) }),
      },
      {
        part: "ppt/slides/slide2.xml",
        shapes: sp({ ph: "title", text: paragraph({ runs: run({ text: "Slide two", size: 3200 }) }) }),
      },
    ],
  });

const deckWithImage = () =>
  buildPptx({
    slides: [
      {
        part: "ppt/slides/slide1.xml",
        shapes: `${pic({ rId: "rId2" })}${sp({ text: paragraph({ runs: run({ text: "with image" }) }) })}`,
        rels: { rId2: { target: "../media/image1.png", type: `${NS.r}/image` } },
      },
    ],
    media: [{ part: "ppt/media/image1.png", bytes: RED_PNG }],
  });

async function waitForReady(container: HTMLElement) {
  await waitFor(() => {
    expect(container.querySelector("[data-testid='rv-root']")?.getAttribute("data-status")).toBe("ready");
  });
}

/** Scope text queries to the main slide area so thumbnails cannot collide. */
function canvasOf(container: HTMLElement) {
  return within(container.querySelector(".radv-canvas") as HTMLElement);
}

describe("DocumentViewer", () => {
  it("reports render issues after the commit instead of during another render", async () => {
    const errors: unknown[][] = [];
    const spy = vi.spyOn(console, "error").mockImplementation((...args: unknown[]) => {
      errors.push(args);
    });
    try {
      // An unimplemented preset shape makes the renderer report an issue mid-render.
      const bytes = buildPptx({
        slides: [
          {
            part: "ppt/slides/slide1.xml",
            shapes: `${sp({ preset: "futureShapeGeom", text: paragraph({ runs: run({ text: "Odd shape" }) }) })}`,
          },
        ],
      });
      const { container } = render(<DocumentViewer source={bytes} debug />);
      await waitForReady(container);
      await waitFor(() => {
        expect(screen.getByText(/Preset shape .* is not implemented/)).toBeTruthy();
      });
      const messages = errors.map((args) => String(args[0]));
      expect(messages.filter((message) => message.includes("while rendering a different component"))).toEqual([]);
    } finally {
      spy.mockRestore();
    }
  });

  it("loads a presentation from bytes and shows the first slide with navigation chrome", async () => {
    const { container } = render(<DocumentViewer source={twoSlides()} />);
    await waitForReady(container);
    expect(canvasOf(container).getByText("Slide one")).toBeTruthy();
    expect(screen.getByText("1 / 2")).toBeTruthy();
    expect(screen.getByLabelText("Next slide")).toBeTruthy();
    expect(screen.getByLabelText("Previous slide")).toBeDisabled();
  });

  it("navigates with buttons, keyboard and the imperative handle", async () => {
    const ref = createRef<DocumentViewerHandle>();
    const onSlideChange = vi.fn();
    const { container } = render(
      <DocumentViewer ref={ref} source={twoSlides()} onSlideChange={onSlideChange} />,
    );
    await waitForReady(container);

    fireEvent.click(screen.getByLabelText("Next slide"));
    await waitFor(() => expect(canvasOf(container).getByText("Slide two")).toBeTruthy());
    expect(onSlideChange).toHaveBeenLastCalledWith(1);
    expect(screen.getByLabelText("Next slide")).toBeDisabled();

    fireEvent.keyDown(screen.getByTestId("rv-root"), { key: "ArrowLeft" });
    await waitFor(() => expect(canvasOf(container).getByText("Slide one")).toBeTruthy());

    // Imperative handle calls update state, so they run inside act().
    await act(async () => ref.current?.next());
    await waitFor(() => expect(canvasOf(container).getByText("Slide two")).toBeTruthy());
    expect(ref.current?.getCurrentSlide()).toBe(1);
    expect(ref.current?.getSlideCount()).toBe(2);
    await act(async () => ref.current?.first());
    await waitFor(() => expect(canvasOf(container).getByText("Slide one")).toBeTruthy());
  });

  it("supports Home/End/PageDown and ignores keys from editable elements", async () => {
    const { container } = render(<DocumentViewer source={twoSlides()} />);
    await waitForReady(container);
    const root = screen.getByTestId("rv-root");

    fireEvent.keyDown(root, { key: "End" });
    await waitFor(() => expect(canvasOf(container).getByText("Slide two")).toBeTruthy());
    fireEvent.keyDown(root, { key: "Home" });
    await waitFor(() => expect(canvasOf(container).getByText("Slide one")).toBeTruthy());

    const input = document.createElement("input");
    root.appendChild(input);
    fireEvent.keyDown(input, { key: "End" });
    expect(canvasOf(container).getByText("Slide one")).toBeTruthy();
  });

  it("zooms in, out and back to fit, reporting the scale", async () => {
    const onZoomChange = vi.fn();
    const ref = createRef<DocumentViewerHandle>();
    const { container } = render(
      <DocumentViewer ref={ref} source={twoSlides()} onZoomChange={onZoomChange} fit="actual" />,
    );
    await waitForReady(container);
    const initial = ref.current?.getZoom() ?? 0;
    expect(initial).toBeCloseTo(1, 5);

    fireEvent.click(screen.getByLabelText("Zoom in"));
    await waitFor(() => expect(ref.current?.getZoom() ?? 0).toBeGreaterThan(1));
    expect(onZoomChange).toHaveBeenCalled();

    fireEvent.click(screen.getByLabelText("Zoom out"));
    await waitFor(() => expect(ref.current?.getZoom()).toBeCloseTo(1, 5));

    fireEvent.click(screen.getByLabelText("Fit to page"));
    await waitFor(() => expect(ref.current?.getZoom() ?? 0).toBeGreaterThan(0));
  });

  it("loads a presentation from a URL, respecting the download limit", async () => {
    const bytes = twoSlides();
    const fetchMock = vi.fn(async () =>
      new Response(bytes as unknown as BodyInit, {
        status: 200,
        headers: { "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
      }),
    );
    vi.stubGlobal("fetch", fetchMock);
    try {
      const { container } = render(<DocumentViewer source="https://example.com/deck.pptx" />);
      await waitForReady(container);
      expect(canvasOf(container).getByText("Slide one")).toBeTruthy();
      expect(fetchMock).toHaveBeenCalledWith(
        "https://example.com/deck.pptx",
        expect.objectContaining({ credentials: "same-origin" }),
      );
    } finally {
      vi.unstubAllGlobals();
    }

    const onError = vi.fn();
    vi.stubGlobal(
      "fetch",
      vi.fn(async () =>
        new Response(bytes as unknown as BodyInit, { status: 200, headers: { "content-length": String(9_000_000) } }),
      ),
    );
    try {
      render(<DocumentViewer source="https://example.com/huge.pptx" maxDownloadBytes={1_000_000} onError={onError} />);
      await waitFor(() => expect(onError).toHaveBeenCalled());
      expect(onError.mock.calls[0]![0].code).toBe("LIMIT_EXCEEDED");
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("presents from the current slide without any chrome", async () => {
    const onPresentationChange = vi.fn();
    const ref = createRef<DocumentViewerHandle>();
    const { container } = render(
      <DocumentViewer
        ref={ref}
        source={twoSlides()}
        showThumbnails
        debug
        onPresentationChange={onPresentationChange}
      />,
    );
    await waitForReady(container);
    const root = screen.getByTestId("rv-root");
    expect(root.getAttribute("data-presenting")).toBe("false");

    // Present from slide 2: the slideshow keeps the current slide.
    await act(async () => ref.current?.next());
    fireEvent.click(screen.getByLabelText("Present from this slide"));
    await waitFor(() => expect(root.getAttribute("data-presenting")).toBe("true"));
    expect(onPresentationChange).toHaveBeenLastCalledWith(true);
    expect(root.className).toContain("radv-root-presenting");

    // Every control is gone, including the diagnostics panel and the sidebar.
    expect(screen.queryByRole("toolbar")).toBeNull();
    expect(container.querySelector(".radv-thumbnails")).toBeNull();
    expect(container.querySelector(".radv-diagnostics")).toBeNull();
    expect(canvasOf(container).getByText("Slide two")).toBeTruthy();

    // Keys work without the toolbar being focused, and a click advances.
    fireEvent.keyDown(document, { key: "ArrowLeft" });
    await waitFor(() => expect(canvasOf(container).getByText("Slide one")).toBeTruthy());
    fireEvent.click(container.querySelector(".radv-canvas") as HTMLElement);
    await waitFor(() => expect(canvasOf(container).getByText("Slide two")).toBeTruthy());
    // Space at the last slide stays put rather than wrapping around.
    fireEvent.keyDown(document, { key: " " });
    await waitFor(() => expect(canvasOf(container).getByText("Slide two")).toBeTruthy());

    // The hint offers an exit for pointer-only users, and Esc always works.
    const hint = screen.getByTestId("rv-present-hint");
    expect(within(hint).getByText(/Esc/)).toBeTruthy();
    fireEvent.keyDown(document, { key: "Escape" });
    await waitFor(() => expect(root.getAttribute("data-presenting")).toBe("false"));
    expect(onPresentationChange).toHaveBeenLastCalledWith(false);
    // Chrome is back, with the sidebar and diagnostics restored.
    expect(screen.getByRole("toolbar")).toBeTruthy();
    expect(container.querySelector(".radv-diagnostics")).not.toBeNull();
    expect(screen.getByLabelText("Next slide")).toBeTruthy();
  });

  it("supports controlled presentation mode and exits when nothing is loaded", async () => {
    const ref = createRef<DocumentViewerHandle>();
    const { container, rerender } = render(
      <DocumentViewer ref={ref} source={twoSlides()} presentation={false} />,
    );
    await waitForReady(container);
    const root = screen.getByTestId("rv-root");

    // The toolbar button cannot start a slideshow while the prop says no.
    await act(async () => ref.current?.togglePresentation());
    expect(root.getAttribute("data-presenting")).toBe("false");
    expect(ref.current?.isPresenting()).toBe(false);

    rerender(<DocumentViewer ref={ref} source={twoSlides()} presentation />);
    await waitFor(() => expect(root.getAttribute("data-presenting")).toBe("true"));
    expect(screen.queryByRole("toolbar")).toBeNull();

    // A source that fails to load must not leave an empty slideshow on screen.
    const onPresentationChange = vi.fn();
    rerender(
      <DocumentViewer
        ref={ref}
        source={new Uint8Array([1, 2, 3, 4])}
        presentation
        onPresentationChange={onPresentationChange}
      />,
    );
    await waitFor(() => expect(root.getAttribute("data-status")).toBe("error"));
    expect(root.getAttribute("data-presenting")).toBe("false");
    await waitFor(() => expect(onPresentationChange).toHaveBeenCalledWith(false));
  });

  it("falls back to an in-page overlay when the Fullscreen API is unavailable", async () => {
    // jsdom has no requestFullscreen, which models a blocked/absent API.
    const { container } = render(<DocumentViewer source={twoSlides()} />);
    await waitForReady(container);
    const root = screen.getByTestId("rv-root");
    expect(root.getAttribute("data-fullscreen")).toBe("false");

    fireEvent.click(screen.getByLabelText("Enter fullscreen"));
    await waitFor(() => expect(root.getAttribute("data-fullscreen")).toBe("overlay"));
    expect(root.className).toContain("radv-root-overlay");
    expect(screen.getByLabelText("Exit fullscreen")).toBeTruthy();

    fireEvent.click(screen.getByLabelText("Exit fullscreen"));
    await waitFor(() => expect(root.getAttribute("data-fullscreen")).toBe("false"));

    // The `f` shortcut uses the same path.
    fireEvent.keyDown(root, { key: "f" });
    await waitFor(() => expect(root.getAttribute("data-fullscreen")).toBe("overlay"));
    fireEvent.keyDown(document, { key: "Escape" });
    await waitFor(() => expect(root.getAttribute("data-fullscreen")).toBe("false"));
  });

  it("renders thumbnails and jumps to the clicked slide", async () => {
    const { container } = render(<DocumentViewer source={twoSlides()} showThumbnails />);
    await waitForReady(container);
    const tabs = screen.getAllByRole("tab");
    expect(tabs).toHaveLength(2);
    fireEvent.click(tabs[1]!);
    await waitFor(() => expect(canvasOf(container).getByText("Slide two")).toBeTruthy());
  });

  it("only mounts thumbnails once they intersect the viewport", async () => {
    const observed: Element[] = [];
    const callbacks: IntersectionObserverCallback[] = [];
    class FakeObserver {
      constructor(callback: IntersectionObserverCallback) {
        callbacks.push(callback);
      }
      observe(target: Element) {
        observed.push(target);
      }
      unobserve() {}
      disconnect() {}
      takeRecords() {
        return [];
      }
      root = null;
      rootMargin = "";
      thresholds = [];
    }
    const original = globalThis.IntersectionObserver;
    (globalThis as { IntersectionObserver?: unknown }).IntersectionObserver = FakeObserver;

    try {
      const { container } = render(<DocumentViewer source={twoSlides()} showThumbnails />);
      await waitForReady(container);
      // Observation happens in an effect, so wait for it rather than racing it.
      await waitFor(() => expect(observed.length).toBe(2));
      // Only the active slide's thumbnail is mounted before anything intersects.
      expect(container.querySelectorAll(".radv-thumb-slot .radv-slide")).toHaveLength(1);
      await act(async () => {
        callbacks[0]!(
          [{ isIntersecting: true, target: observed[1]! } as unknown as IntersectionObserverEntry],
          {} as IntersectionObserver,
        );
      });
      await waitFor(() =>
        expect(container.querySelectorAll(".radv-thumb-slot .radv-slide")).toHaveLength(2),
      );
    } finally {
      (globalThis as { IntersectionObserver?: unknown }).IntersectionObserver = original;
    }
  });

  it("reports a clear error for a file that is not a presentation", async () => {
    const onError = vi.fn();
    const junk = new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]);
    const { container } = render(<DocumentViewer source={junk} onError={onError} />);
    await waitFor(() => {
      expect(container.querySelector("[data-testid='rv-error-code']")?.textContent).toBe("NOT_A_PRESENTATION");
    });
    expect(onError).toHaveBeenCalledTimes(1);
    expect(onError.mock.calls[0]?.[0]?.code).toBe("NOT_A_PRESENTATION");
  });

  it("detects an encrypted or legacy binary file", async () => {
    const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]);
    const { container } = render(<DocumentViewer source={ole} />);
    await waitFor(() => {
      expect(container.querySelector("[data-testid='rv-error-code']")?.textContent).toBe("ENCRYPTED");
    });
  });

  it("invokes onLoad with presentation metadata and onDiagnostics", async () => {
    const onLoad = vi.fn();
    const { container } = render(<DocumentViewer source={twoSlides()} onLoad={onLoad} />);
    await waitForReady(container);
    expect(onLoad).toHaveBeenCalledTimes(1);
    const info = onLoad.mock.calls[0]?.[0];
    expect(info.slideCount).toBe(2);
    expect(info.widthEmu).toBe(12_192_000);
    expect(info.aspectRatio).toBeCloseTo(12_192_000 / 6_858_000, 5);
  });

  it("shows a diagnostics panel in debug mode", async () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: pic({ rId: "rId2" }),
          rels: { rId2: { target: "../media/missing.png", type: `${NS.r}/image` } },
        },
      ],
    });
    const { container } = render(<DocumentViewer source={bytes} debug />);
    await waitForReady(container);
    const panel = screen.getByTestId("rv-diagnostics");
    expect(panel.textContent).toContain("picture.missing-part");
    fireEvent.click(within(panel).getByLabelText("Close diagnostics"));
    expect(screen.queryByTestId("rv-diagnostics")).toBeNull();
  });

  it("explains an unmeasured container instead of rendering nothing", async () => {
    setZeroSize(true);
    const { container } = render(<DocumentViewer source={twoSlides()} />);
    await waitForReady(container);
    expect(container.querySelector(".radv-notice")?.textContent).toContain("usable size");
  });

  it("keeps the slide aspect ratio and warns about an empty slide", async () => {
    const bytes = buildPptx({
      slides: [
        { part: "ppt/slides/slide1.xml", shapes: "" },
        {
          part: "ppt/slides/slide2.xml",
          shapes: sp({ ph: "title", text: paragraph({ runs: run({ text: "Has content" }) }) }),
        },
      ],
    });
    const { container } = render(<DocumentViewer source={bytes} />);
    await waitForReady(container);
    expect(container.querySelector(".radv-empty-slide")).not.toBeNull();
    const inner = container.querySelector(".radv-canvas-inner") as HTMLElement;
    const width = Number.parseFloat(inner.style.width);
    const height = Number.parseFloat(inner.style.height);
    expect(width / height).toBeCloseTo(12_192_000 / 6_858_000, 2);
  });

  it("contains a slide render failure and reports it instead of going blank", async () => {
    const parsedBytes = deckWithImage();
    const { parsePresentationBytes } = await import("../../src/parser/presentation");
    const parsed = parsePresentationBytes(parsedBytes);
    const slide = parsed.model.slides[0]!;

    // A media store that throws reproduces a real rendering fault inside the
    // picture layer; the viewer must contain it and report it per slide.
    const brokenMedia = {
      kind: () => "png" as const,
      url: () => {
        throw new Error("decode failed");
      },
      mime: () => "image/png",
      byteLength: () => 0,
      dispose: () => undefined,
      activeUrls: 0,
    };

    const { container } = render(
      <SlideCanvasForTest
        model={parsed.model}
        media={brokenMedia}
        slide={slide}
        slideWidth={1280}
        slideHeight={720}
      />,
    );
    const panel = await screen.findByTestId("rv-slide-error");
    expect(panel.textContent).toContain("decode failed");
    expect(container).toBeTruthy();
  });

  it("supports custom loader and error renderers", async () => {
    const junk = new Uint8Array([9, 9, 9, 9]);
    const { container } = render(
      <DocumentViewer
        source={junk}
        renderLoader={() => <div data-testid="custom-loader">loading…</div>}
        renderError={(error, retry) => (
          <div data-testid="custom-error">
            {error.code}
            <button type="button" onClick={retry}>
              retry
            </button>
          </div>
        )}
      />,
    );
    await waitFor(() => expect(screen.getByTestId("custom-error")).toBeTruthy());
    expect(screen.getByTestId("custom-error").textContent).toContain("NOT_A_PRESENTATION");
    fireEvent.click(screen.getByText("retry"));
    await waitFor(() => expect(screen.getByTestId("custom-error")).toBeTruthy());
  });

  it("revokes object URLs when it unmounts", async () => {
    const revoke = vi.spyOn(URL, "revokeObjectURL");
    const { unmount, container } = render(<DocumentViewer source={deckWithImage()} />);
    await waitForReady(container);
    const image = container.querySelector("img");
    await waitFor(() => expect(image?.getAttribute("src") ?? "").toMatch(/^(blob:|data:)/));
    const url = image?.getAttribute("src") ?? "";
    unmount();
    expect(revoke).toHaveBeenCalledWith(url);
  });

  it("never shows a stale slide when the source changes rapidly", async () => {
    const slow = twoSlides();
    const fast = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({ ph: "title", text: paragraph({ runs: run({ text: "Fresh deck" }) }) }),
        },
      ],
    });
    const { container, rerender } = render(<DocumentViewer source={slow} />);
    rerender(<DocumentViewer source={fast} />);
    await waitFor(() => expect(canvasOf(container).getByText("Fresh deck")).toBeTruthy());
    expect(canvasOf(container).queryByText("Slide one")).toBeNull();
  });

  it("renders safely inside StrictMode and still cleans up", async () => {
    const revoke = vi.spyOn(URL, "revokeObjectURL");
    const { container, unmount } = render(
      <StrictMode>
        <DocumentViewer source={deckWithImage()} />
      </StrictMode>,
    );
    await waitForReady(container);
    await waitFor(() => {
      const src = container.querySelector("img")?.getAttribute("src") ?? "";
      expect(src).toMatch(/^(blob:|data:)/);
    });
    expect(container.querySelector(".radv-root")).not.toBeNull();
    unmount();
    expect(revoke).toHaveBeenCalled();
  });

  it("respects a controlled slide index", async () => {
    const bytes = twoSlides();
    const { container, rerender } = render(<DocumentViewer source={bytes} slide={0} />);
    await waitForReady(container);
    expect(canvasOf(container).getByText("Slide one")).toBeTruthy();
    rerender(<DocumentViewer source={bytes} slide={1} />);
    await waitFor(() => expect(canvasOf(container).getByText("Slide two")).toBeTruthy());
  });

  it("hides the toolbar when controls are disabled", async () => {
    const { container } = render(<DocumentViewer source={twoSlides()} showControls={false} />);
    await waitForReady(container);
    expect(container.querySelector(".radv-toolbar")).toBeNull();
    expect(canvasOf(container).getByText("Slide one")).toBeTruthy();
  });

  it("localises the toolbar through the labels prop", async () => {
    const { container } = render(
      <DocumentViewer
        source={twoSlides()}
        labels={{ next: "Slajd dalej", slideOf: (current, total) => `${current} z ${total}` }}
      />,
    );
    await waitForReady(container);
    expect(screen.getByLabelText("Slajd dalej")).toBeTruthy();
    expect(screen.getByText("1 z 2")).toBeTruthy();
  });

  it("renders nothing when no source is provided", () => {
    const { container } = render(<DocumentViewer />);
    expect(container.querySelector("[data-testid='rv-root']")?.getAttribute("data-status")).toBe("idle");
    expect(container.querySelector(".radv-slide")).toBeNull();
  });
});

describe("DocumentViewer with a Word document", () => {
  const twoPageDocx = () =>
    buildDocx({
      body: `${docxParagraph("First page of the document")}${wParagraph({
        runs: wRun("Second page"),
        pageBreakBefore: true,
      })}`,
    });

  it("loads a .docx, reports its format and renders page text", async () => {
    const onLoad = vi.fn();
    const { container } = render(<DocumentViewer source={twoPageDocx()} showThumbnails onLoad={onLoad} />);
    await waitForReady(container);

    expect(canvasOf(container).getByText("First page of the document")).toBeTruthy();
    expect(container.querySelectorAll(".radv-thumb-item")).toHaveLength(2);
    expect(onLoad).toHaveBeenCalledWith(
      expect.objectContaining({ format: "docx", slideCount: 2 }),
    );
    // A Word page is US Letter, not a 16:9 stage.
    expect(onLoad.mock.calls[0]?.[0]?.aspectRatio).toBeCloseTo(8.5 / 11, 3);
  });

  it("paginates with the toolbar and the keyboard like a deck", async () => {
    const { container } = render(<DocumentViewer source={twoPageDocx()} />);
    await waitForReady(container);
    expect(screen.getByText("1 / 2")).toBeTruthy();

    await act(async () => {
      fireEvent.click(screen.getByLabelText("Next slide"));
    });
    expect(canvasOf(container).getByText("Second page")).toBeTruthy();
    expect(screen.getByText("2 / 2")).toBeTruthy();
  });

  it("exposes the format on the load metadata for presentations too", async () => {
    const onLoad = vi.fn();
    const { container } = render(<DocumentViewer source={twoSlides()} onLoad={onLoad} />);
    await waitForReady(container);
    expect(onLoad).toHaveBeenCalledWith(expect.objectContaining({ format: "pptx" }));
  });
});

describe("page renderer seam", () => {
  it("lets a format draw its own pages through the canvas", () => {
    const slide = {
      index: 0,
      partName: "pdf#page1",
      name: null,
      slideshowHidden: false,
      shapes: [],
      inheritedShapes: [],
      background: null,
      diagnostics: [],
    };
    const renderPage = vi.fn((_context: Parameters<PageRenderer>[0]) => (
      <div data-testid="custom-page">custom page</div>
    ));
    const { getByTestId } = render(
      <SlideCanvas
        model={{ slides: [slide] } as never}
        media={createEmptyMediaStore()}
        slide={slide as never}
        slideIndex={0}
        slideWidth={612}
        slideHeight={792}
        scale={1.5}
        size={{ width: 900, height: 600, measured: true }}
        containerRef={() => undefined}
        labels={mergeLabels(undefined)}
        renderPage={renderPage}
      />,
    );
    expect(getByTestId("custom-page")).toBeTruthy();
    expect(renderPage).toHaveBeenCalledWith(
      expect.objectContaining({ width: 612, height: 792, scale: 1.5, interactive: true }),
    );
    // The slide view is not used when the format draws its own pages.
    expect(renderPage.mock.calls[0]?.[0]?.page).toBe(slide);
  });
});

/** Harness that renders `SlideCanvas` with a stubbed media store. */
function SlideCanvasForTest({
  model,
  media,
  slide,
  slideWidth,
  slideHeight,
}: {
  model: Parameters<typeof SlideCanvas>[0]["model"];
  media: Parameters<typeof SlideCanvas>[0]["media"];
  slide: Parameters<typeof SlideCanvas>[0]["slide"];
  slideWidth: number;
  slideHeight: number;
}) {
  return (
    <SlideCanvas
      model={model}
      media={media}
      slide={slide}
      slideIndex={slide?.index ?? 0}
      slideWidth={slideWidth}
      slideHeight={slideHeight}
      scale={1}
      size={{ width: 900, height: 600, measured: true }}
      containerRef={() => undefined}
      labels={mergeLabels(undefined)}
    />
  );
}
