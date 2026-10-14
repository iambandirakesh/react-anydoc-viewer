import { expect, test, type Page } from "@playwright/test";

/**
 * End-to-end checks against the demo application in real Chrome.
 *
 * The demo generates an eight-slide deck in the browser (`examples/demo/src/
 * sampleDeck.ts`) covering text, shapes, images, a group, a table, charts, an
 * empty slide and deliberately unsupported content, so the whole viewer can be
 * exercised without shipping presentation files.
 */

const SAMPLE_SLIDES = 8;

async function loadSampleDeck(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByRole("button", { name: "Load sample deck" }).first().click();
  await expect(page.locator(".radv-root")).toHaveAttribute("data-status", "ready");
  await expect(page.locator(".radv-counter")).toHaveText(`1 / ${SAMPLE_SLIDES}`);
}

const counter = (page: Page) => page.locator(".radv-counter");
const zoomLabel = (page: Page) => page.locator(".radv-zoom");
const toolbarButton = (page: Page, name: string) => page.locator(`.radv-toolbar button[aria-label="${name}"]`);
const zoomMenu = (page: Page) => page.getByRole("menu", { name: "Zoom and fit" });
const zoomTrigger = (page: Page) => page.getByRole("button", { name: "Zoom and fit" });

/**
 * The chrome hides itself until the viewer is hovered (or one of its controls has
 * focus), so the pointer has to come in before anything in the pill can be used.
 */
async function revealChrome(page: Page): Promise<void> {
  await page.locator(".radv-canvas").hover();
  await expect(page.locator(".radv-toolbar")).toHaveCSS("opacity", "1");
}

async function clickToolbar(page: Page, name: string): Promise<void> {
  await revealChrome(page);
  await toolbarButton(page, name).click();
}

/** Zoom stops and fit modes share one dropdown above the percentage. */
async function openZoomMenu(page: Page): Promise<void> {
  await revealChrome(page);
  await zoomTrigger(page).click();
  await expect(zoomMenu(page)).toBeVisible();
}

async function chooseFromZoomMenu(page: Page, name: string): Promise<void> {
  await openZoomMenu(page);
  await page.getByRole("menuitemradio", { name, exact: true }).click();
}

test.describe("viewer chrome", () => {
  test("loads a deck and reports its metadata", async ({ page }) => {
    await loadSampleDeck(page);
    await expect(page.locator(".radv-thumb-item")).toHaveCount(SAMPLE_SLIDES);
    await expect(page.getByText("13.33in × 7.50in")).toBeVisible();
    // The sample deck has five slides' worth of diagnostics (unsupported frames,
    // approximated table styling, an unimplemented preset).
    await expect(page.locator(".demo__stats dd").last()).toHaveText("none");
  });

  test("navigates with the toolbar, thumbnails and the keyboard", async ({ page }) => {
    await loadSampleDeck(page);

    await revealChrome(page);
    await expect(toolbarButton(page, "Previous slide")).toBeDisabled();

    await clickToolbar(page, "Next slide");
    await expect(counter(page)).toHaveText(`2 / ${SAMPLE_SLIDES}`);

    // The pill is deliberately short: the jumps to the ends stay a keyboard and
    // handle action rather than two more buttons in every reader's view.
    await page.locator(".radv-canvas").click();
    await page.keyboard.press("End");
    await expect(counter(page)).toHaveText(`${SAMPLE_SLIDES} / ${SAMPLE_SLIDES}`);
    await expect(toolbarButton(page, "Next slide")).toBeDisabled();

    await clickToolbar(page, "Previous slide");
    await expect(counter(page)).toHaveText(`${SAMPLE_SLIDES - 1} / ${SAMPLE_SLIDES}`);

    await page.locator(".radv-canvas").click();
    await page.keyboard.press("Home");
    await expect(counter(page)).toHaveText(`1 / ${SAMPLE_SLIDES}`);

    // Thumbnails navigate too, and the active item follows.
    await page.getByRole("tab", { name: "Slide 4" }).click();
    await expect(counter(page)).toHaveText(`4 / ${SAMPLE_SLIDES}`);
    await expect(page.locator(".radv-thumb-item").nth(3)).toHaveClass(/radv-thumb-active/);

    // Keyboard shortcuts (the viewer itself must have focus).
    await page.locator(".radv-canvas").click();
    await page.keyboard.press("ArrowRight");
    await expect(counter(page)).toHaveText(`5 / ${SAMPLE_SLIDES}`);
    await page.keyboard.press("Home");
    await expect(counter(page)).toHaveText(`1 / ${SAMPLE_SLIDES}`);
    await page.keyboard.press("End");
    await expect(counter(page)).toHaveText(`${SAMPLE_SLIDES} / ${SAMPLE_SLIDES}`);
    await page.keyboard.press("ArrowLeft");
    await expect(counter(page)).toHaveText(`${SAMPLE_SLIDES - 1} / ${SAMPLE_SLIDES}`);
  });

  test("floats the chrome as a bottom pill that only shows on hover", async ({ page }) => {
    await loadSampleDeck(page);

    // Nothing is in the way while reading: the pill is invisible and inert until
    // the viewer is hovered. (It is not `display: none`, so focus can still reach
    // it — the keyboard path is exercised further down.)
    const toolbar = page.locator(".radv-toolbar");
    await expect(toolbar).toHaveCSS("opacity", "0");
    await expect(toolbar).toHaveCSS("pointer-events", "none");
    await revealChrome(page);
    await expect(toolbar).toHaveCSS("pointer-events", "auto");
    await page.mouse.move(4, 4);
    await expect(toolbar).toHaveCSS("opacity", "0");

    const pill = await toolbar.boundingBox();
    const stage = await page.locator(".radv-stage").boundingBox();
    const slide = await page.locator(".radv-canvas-inner").boundingBox();
    expect(pill).not.toBeNull();

    // A pill, not a header: it hugs the bottom of the stage, centred on it, and
    // stays small — a third of the stage is the budget, not a full-width bar.
    expect(pill!.y).toBeGreaterThan(stage!.y + stage!.height / 2);
    expect(pill!.height).toBeLessThan(48);
    expect(pill!.width).toBeLessThan(stage!.width * 0.5);
    expect(Math.abs(pill!.x + pill!.width / 2 - (stage!.x + stage!.width / 2))).toBeLessThan(4);
    // Fit-to-page leaves room for it rather than hiding the slide's bottom behind it.
    expect(slide!.y + slide!.height).toBeLessThanOrEqual(pill!.y);

    // The dropdown opens on click, not on hover: the pointer may cross the pill
    // while reading without covering the slide with a menu.
    await revealChrome(page);
    await zoomTrigger(page).hover();
    await expect(zoomMenu(page)).toHaveCount(0);
    await expect(zoomTrigger(page)).toHaveAttribute("aria-expanded", "false");
    await zoomTrigger(page).click();
    await expect(zoomMenu(page)).toBeVisible();
    // Eleven zoom stops and the four fit modes, divided into two groups.
    await expect(zoomMenu(page).getByRole("menuitemradio")).toHaveCount(15);
    await expect(zoomMenu(page).getByRole("separator")).toHaveCount(1);

    // The trigger is one boxed control: a percentage half, a divider, a caret.
    const trigger = await zoomTrigger(page).boundingBox();
    const caret = await page.locator(".radv-caret").boundingBox();
    expect(trigger).not.toBeNull();
    expect(caret!.x).toBeGreaterThan(trigger!.x + trigger!.width / 2);
    expect(caret!.height).toBeGreaterThan(trigger!.height - 4);

    // Clicking away dismisses it again.
    await page.mouse.click(stage!.x + 20, stage!.y + 20);
    await expect(zoomMenu(page)).toHaveCount(0);

    // The same menu is reachable without a pointer: Esc closes and hands focus
    // back, ArrowDown opens it on the choice in effect.
    await openZoomMenu(page);
    await page.keyboard.press("Escape");
    await expect(zoomMenu(page)).toHaveCount(0);
    await expect(zoomTrigger(page)).toBeFocused();

    await page.keyboard.press("ArrowDown");
    await expect(zoomMenu(page)).toBeVisible();
    await expect(page.getByRole("menuitemradio", { name: "Fit to page", exact: true })).toBeFocused();
    await page.keyboard.press("Escape");
    await expect(zoomMenu(page)).toHaveCount(0);
  });

  test("zooms and switches fit mode", async ({ page }) => {
    await loadSampleDeck(page);

    const initial = await zoomLabel(page).textContent();
    await clickToolbar(page, "Zoom in");
    await expect(zoomLabel(page)).not.toHaveText(initial ?? "");

    await chooseFromZoomMenu(page, "100%");
    await expect(zoomLabel(page)).toHaveText("100%");
    await openZoomMenu(page);
    await expect(page.getByRole("menuitemradio", { name: "100%", exact: true })).toHaveAttribute(
      "aria-checked",
      "true",
    );
    await page.keyboard.press("Escape");
    await expect(zoomMenu(page)).toHaveCount(0);

    await chooseFromZoomMenu(page, "Fit to width");
    await expect(zoomLabel(page)).not.toHaveText("100%");
    await openZoomMenu(page);
    await expect(page.getByRole("menuitemradio", { name: "Fit to width", exact: true })).toHaveAttribute(
      "aria-checked",
      "true",
    );
    await page.keyboard.press("Escape");

    await chooseFromZoomMenu(page, "Fit to page");
    await openZoomMenu(page);
    await expect(page.getByRole("menuitemradio", { name: "Fit to page", exact: true })).toHaveAttribute(
      "aria-checked",
      "true",
    );
    await page.keyboard.press("Escape");

    await page.keyboard.press("-");
    const zoomedOut = Number.parseInt((await zoomLabel(page).textContent()) ?? "0", 10);
    expect(zoomedOut).toBeLessThan(100);
    await page.keyboard.press("+");
    expect(Number.parseInt((await zoomLabel(page).textContent()) ?? "0", 10)).toBeGreaterThan(zoomedOut);
  });

  test("toggles the thumbnail sidebar", async ({ page }) => {
    await loadSampleDeck(page);
    await expect(page.locator(".radv-thumbnails")).toBeVisible();

    await clickToolbar(page, "Toggle page thumbnails");
    await expect(page.locator(".radv-thumbnails")).toHaveCount(0);

    await clickToolbar(page, "Toggle page thumbnails");
    await expect(page.locator(".radv-thumbnails")).toBeVisible();
  });

  test("enters and leaves fullscreen, natively or through the overlay", async ({ page }) => {
    await loadSampleDeck(page);
    const root = page.locator(".radv-root");

    await clickToolbar(page, "Enter fullscreen");
    // Either the browser entered fullscreen or the viewer expanded itself.
    await expect(root).not.toHaveAttribute("data-fullscreen", "false");
    const mode = await root.getAttribute("data-fullscreen");
    expect(["native", "overlay"]).toContain(mode);
    await expect(toolbarButton(page, "Exit fullscreen")).toBeVisible();

    if (mode === "overlay") {
      // The overlay covers the viewport and Esc is the documented way out.
      const box = await root.boundingBox();
      expect(box?.width).toBeGreaterThan(1000);
      await page.keyboard.press("Escape");
    } else {
      // Chrome's own Escape handling needs a real key event, which synthetic
      // key presses are not, so the toolbar button is used instead.
      await clickToolbar(page, "Exit fullscreen");
    }
    await expect(root).toHaveAttribute("data-fullscreen", "false");
    await expect(page.locator(".radv-thumbnails")).toBeVisible();
  });

  test("presents the deck without chrome and advances on click, Space and arrows", async ({ page }) => {
    await loadSampleDeck(page);
    const root = page.locator(".radv-root");
    const hint = page.getByTestId("rv-present-hint");

    await clickToolbar(page, "Present from this slide");
    await expect(root).toHaveAttribute("data-presenting", "true");
    await expect(page.getByRole("toolbar")).toHaveCount(0);
    await expect(page.locator(".radv-thumbnails")).toHaveCount(0);
    await expect(page.locator(".radv-counter")).toHaveCount(0);
    await expect(hint).toContainText(`1 / ${SAMPLE_SLIDES}`);

    // The hint is a single compact row: it must not wrap into a tall blob.
    const hintBox = await hint.boundingBox();
    expect(hintBox!.height).toBeLessThan(40);
    expect(hintBox!.width).toBeLessThan(520);

    await page.keyboard.press("Space");
    await expect(hint).toContainText(`2 / ${SAMPLE_SLIDES}`);
    await page.keyboard.press("ArrowRight");
    await expect(hint).toContainText(`3 / ${SAMPLE_SLIDES}`);
    await page.locator(".radv-canvas").click();
    await expect(hint).toContainText(`4 / ${SAMPLE_SLIDES}`);
    await page.keyboard.press("Home");
    await expect(hint).toContainText(`1 / ${SAMPLE_SLIDES}`);

    // Leave through the hint's button (pointer-only users) and confirm the chrome
    // comes back exactly once.
    await page.mouse.move(200, 300);
    await expect(hint).toBeVisible();
    await page.getByRole("button", { name: "Exit presentation" }).click();
    await expect(root).toHaveAttribute("data-presenting", "false");
    await expect(page.getByRole("toolbar")).toBeVisible();

    // Esc is the other documented way out.
    await clickToolbar(page, "Present from this slide");
    await expect(root).toHaveAttribute("data-presenting", "true");
    await page.keyboard.press("Escape");
    await expect(root).toHaveAttribute("data-presenting", "false");
    await expect(page.getByRole("toolbar")).toBeVisible();
  });

  test("the demo drives presentation mode through its own control", async ({ page }) => {
    await page.goto("/");
    await page.getByRole("button", { name: "Load sample deck" }).first().click();
    await expect(page.locator(".radv-root")).toHaveAttribute("data-status", "ready");

    const toggle = page.getByRole("checkbox", { name: "Presentation mode" });
    await toggle.check();
    await expect(page.locator(".radv-root")).toHaveAttribute("data-presenting", "true");

    // Esc reports the change back, so the controlled checkbox unticks itself.
    await page.keyboard.press("Escape");
    await expect(page.locator(".radv-root")).toHaveAttribute("data-presenting", "false");
    await expect(toggle).not.toBeChecked();
    await expect(page.locator('.demo__stats dt:has-text("Presenting") + dd')).toHaveText("no");
  });

  test("shows diagnostics for unsupported content in debug mode", async ({ page }) => {
    await page.goto("/");
    await page.getByRole("checkbox", { name: /Debug mode/ }).check();
    await page.getByRole("button", { name: "Load sample deck" }).first().click();
    await expect(page.locator(".radv-root")).toHaveAttribute("data-status", "ready");

    const panel = page.locator(".radv-diagnostics");
    await expect(panel).toBeVisible();
    await expect(panel).toContainText("ole.unsupported");
    await expect(panel).toContainText("smartart.unsupported");

    // Slide 7 is the unsupported-content slide: every element is labelled.
    await page.getByRole("tab", { name: "Slide 7" }).click();
    await expect(page.locator('.radv-canvas [role="note"]')).toHaveCount(4);
  });

  test("distinguishes an empty slide from a failed render", async ({ page }) => {
    await loadSampleDeck(page);
    await page.getByRole("tab", { name: "Slide 8" }).click();
    await expect(page.locator(".radv-canvas .radv-empty-slide")).toHaveText(/empty/i);
    // The thumbnail must not repeat the notice.
    await expect(page.locator(".radv-thumbnails .radv-empty-slide")).toHaveCount(0);
    await expect(page.locator('[data-testid="rv-slide-error"]')).toHaveCount(0);
  });

  test("renders slide content as real geometry and text", async ({ page }) => {
    await loadSampleDeck(page);

    // Slide 3: shapes become SVG paths.
    await page.getByRole("tab", { name: "Slide 3" }).click();
    await expect(page.locator(".radv-canvas svg path").first()).toBeVisible();
    await expect(page.locator(".radv-canvas")).toContainText("Shapes, fills and outlines");

    // Slide 4: pictures and a group.
    await page.getByRole("tab", { name: "Slide 4" }).click();
    await expect(page.locator(".radv-canvas img")).toHaveCount(3);
    await expect(page.locator(".radv-canvas")).toContainText("Grouped shapes");

    // Slide 5: a table plus a sanitised hyperlink.
    await page.getByRole("tab", { name: "Slide 5" }).click();
    await expect(page.locator(".radv-canvas")).toContainText("Merged cell spanning three columns");
    await expect(page.locator('.radv-canvas a[href="https://example.com/"]')).toHaveCount(1);

    // Slide 6: charts, including a pie that must not have a hole.
    await page.getByRole("tab", { name: "Slide 6" }).click();
    await expect(page.locator(".radv-canvas")).toContainText("Downloads per month");
    const pieArcs = await page.locator(".radv-canvas svg path").evaluateAll((paths) =>
      paths.map((path) => (path.getAttribute("d") ?? "").split("A").length - 1),
    );
    expect(pieArcs.filter((arcs) => arcs === 1).length).toBeGreaterThanOrEqual(3);
  });

  test("keeps the slide aspect ratio while the window is resized", async ({ page }) => {
    await loadSampleDeck(page);
    const slide = page.locator(".radv-canvas-inner");
    const size = async () => {
      const box = await slide.boundingBox();
      return { width: box?.width ?? 0, height: box?.height ?? 0 };
    };

    // Re-fitting happens in a ResizeObserver callback, so poll rather than reading
    // the box immediately after the viewport changes.
    await page.setViewportSize({ width: 900, height: 800 });
    await expect.poll(async () => Math.round((await size()).height)).toBeGreaterThan(0);
    const narrow = await size();

    await page.setViewportSize({ width: 1600, height: 900 });
    await expect.poll(async () => Math.round((await size()).width)).toBeGreaterThan(narrow.width + 1);
    const wide = await size();

    expect(wide.width).toBeGreaterThan(narrow.width);
    // 16:9 in both cases, within a pixel of rounding.
    expect(Math.abs(narrow.width / narrow.height - 16 / 9)).toBeLessThan(0.02);
    expect(Math.abs(wide.width / wide.height - 16 / 9)).toBeLessThan(0.02);
  });

  test("loads a user-provided file and reports a broken one", async ({ page }) => {
    await page.goto("/");
    const input = page.locator('input[type="file"]');

    // Not a ZIP at all: the viewer must explain itself instead of going blank.
    await input.setInputFiles({
      name: "broken.pptx",
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      buffer: Buffer.from("this is not a presentation"),
    });
    await expect(page.locator(".radv-root")).toHaveAttribute("data-status", "error");
    await expect(page.locator(".radv-status-error")).toBeVisible();
    await expect(page.locator(".demo__stats dd").last()).toContainText("NOT_A_PRESENTATION");
  });

  test("produces no console errors while browsing every slide", async ({ page }) => {
    const errors: string[] = [];
    page.on("console", (message) => {
      if (message.type() === "error") errors.push(message.text());
    });
    page.on("pageerror", (error) => errors.push(error.message));

    await loadSampleDeck(page);
    for (let index = 1; index <= SAMPLE_SLIDES; index += 1) {
      await page.getByRole("tab", { name: `Slide ${index}` }).click();
      await expect(counter(page)).toHaveText(`${index} / ${SAMPLE_SLIDES}`);
    }
    expect(errors).toEqual([]);
  });
});

/** Load a generated document from the demo's own toolbar. */
async function loadSample(page: Page, name: string): Promise<void> {
  await page.goto("/");
  await page.getByRole("button", { name }).first().click();
  await expect(page.locator(".radv-root")).toHaveAttribute("data-status", "ready");
}

test.describe("Word documents", () => {
  test("lays a .docx out into pages in a real browser", async ({ page }) => {
    const errors: string[] = [];
    page.on("console", (message) => {
      if (message.type() === "error") errors.push(message.text());
    });
    page.on("pageerror", (error) => errors.push(error.message));

    await loadSample(page, "Load sample .docx");
    await page.getByRole("checkbox", { name: /Debug mode/ }).check();

    await expect(page.locator('.demo__stats dt:has-text("Format") + dd')).toHaveText("Word document");
    await expect(page.locator(".radv-canvas")).toContainText("A Word document, rendered in the browser");

    // Real height measurement, real pagination: neither the estimator nor the
    // overflow guard may have been needed.
    const panel = page.locator(".radv-diagnostics");
    await expect(panel).toBeVisible();
    await expect(panel).not.toContainText("docx.layout.estimated");
    await expect(panel).not.toContainText("docx.page.overflow");
    // Headers really are reported rather than silently dropped.
    await expect(panel).toContainText("docx.header-footer");

    // The table, the picture and the lists all made it onto the page.
    await expect(page.locator(".radv-canvas [data-shape-kind='table']")).toHaveCount(1);
    await expect(page.locator(".radv-canvas img")).toHaveCount(1);
    await expect(page.locator(".radv-canvas")).toContainText("pdf.js, lazily loaded");

    // Blocks are stacked without overlapping and stay inside the page.
    const layout = await page.locator(".radv-canvas .radv-slide").evaluate((slide) => {
      const page = slide.getBoundingClientRect();
      const blocks = [...slide.querySelectorAll<HTMLElement>("[data-shape-name]")]
        .map((element) => element.getBoundingClientRect())
        .sort((a, b) => a.top - b.top);
      let overlapping = 0;
      for (let index = 1; index < blocks.length; index += 1) {
        if (blocks[index]!.top < blocks[index - 1]!.bottom - 1) overlapping += 1;
      }
      const outside = blocks.filter(
        (block) => block.top < page.top - 1 || block.bottom > page.bottom + 1,
      ).length;
      return { count: blocks.length, overlapping, outside };
    });
    expect(layout.count).toBeGreaterThan(3);
    expect(layout.overlapping).toBe(0);
    expect(layout.outside).toBe(0);

    // A second page exists (the explicit break) and navigation reaches it.
    const pages = Number(await page.locator('.demo__stats dt:has-text("Slides / pages") + dd').textContent());
    expect(pages).toBeGreaterThanOrEqual(2);
    await page.locator(".radv-canvas").click();
    await page.keyboard.press("End");
    await expect(counter(page)).toHaveText(`${pages} / ${pages}`);
    await expect(page.locator(".radv-canvas")).toContainText("Appendix");
    expect(errors).toEqual([]);
  });
});

test.describe("PDF documents", () => {
  test("rasterises each .pdf page onto a canvas", async ({ page }) => {
    const errors: string[] = [];
    page.on("console", (message) => {
      if (message.type() === "error") errors.push(message.text());
    });
    page.on("pageerror", (error) => errors.push(error.message));

    await loadSample(page, "Load sample .pdf");
    await expect(page.locator('.demo__stats dt:has-text("Format") + dd')).toHaveText("PDF document");
    await expect(page.locator('.demo__stats dt:has-text("Slides / pages") + dd')).toHaveText("3");
    await expect(counter(page)).toHaveText("1 / 3");

    // Thumbnails have their own canvases, so the main page is scoped to the stage.
    const canvas = page.locator('.radv-canvas [data-testid="rv-pdf-page"]');
    await expect(canvas).toBeVisible();

    // Real pixels, not an empty canvas: the page has ink at the display size, so
    // it stays sharp instead of being upscaled from a thumbnail-sized bitmap.
    const first = await canvas.evaluate((element) => {
      const target = element as HTMLCanvasElement;
      const context = target.getContext("2d");
      const box = target.getBoundingClientRect();
      if (!context) return { ink: 0, width: target.width, height: target.height, boxWidth: box.width };
      const data = context.getImageData(0, 0, target.width, target.height).data;
      let ink = 0;
      for (let index = 0; index < data.length; index += 4) {
        if (data[index]! < 200) ink += 1;
      }
      return { ink, width: target.width, height: target.height, boxWidth: box.width };
    });
    expect(first.ink).toBeGreaterThan(100);
    // US Letter: 8.5in of page, backing store at least one device pixel per CSS px.
    expect(first.width).toBeGreaterThanOrEqual(Math.floor(first.boxWidth));
    expect(first.boxWidth).toBeGreaterThan(300);

    // Zooming re-rasterises instead of stretching the bitmap.
    await chooseFromZoomMenu(page, "100%");
    await clickToolbar(page, "Zoom in");
    await expect
      .poll(async () => (await canvas.evaluate((element) => (element as HTMLCanvasElement).width)) ?? 0)
      .toBeGreaterThan(first.width);
    await chooseFromZoomMenu(page, "Fit to page");

    // Pages navigate like slides, including their thumbnails.
    await expect(page.locator(".radv-thumb-item")).toHaveCount(3);
    await page.getByRole("tab", { name: "Slide 2" }).click();
    await expect(counter(page)).toHaveText("2 / 3");

    // Presentation mode shows the rasterised page and advances with the keyboard.
    await clickToolbar(page, "Present from this slide");
    await expect(page.locator(".radv-root")).toHaveAttribute("data-presenting", "true");
    await expect(page.locator('[data-testid="rv-pdf-page"]')).toBeVisible();
    await page.keyboard.press("Space");
    await expect(page.locator("[data-testid='rv-present-hint']")).toContainText("3 / 3");
    await page.keyboard.press("Escape");

    expect(errors).toEqual([]);
  });

  test("explains a file that only claims to be a PDF", async ({ page }) => {
    await page.goto("/");
    await page.locator('input[type="file"]').setInputFiles({
      name: "broken.pdf",
      mimeType: "application/pdf",
      buffer: Buffer.from("%PDF-1.4\nnot really a pdf\n%%EOF\n"),
    });
    await expect(page.locator(".radv-root")).toHaveAttribute("data-status", "error");
    await expect(page.locator(".demo__stats dd").last()).toContainText("MALFORMED_DOCUMENT");
  });
});
