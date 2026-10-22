/**
 * Load-time benchmark for the three heavy formats.
 *
 * The clock is read *in the page*, not in Node: `performance.now()` is stamped by
 * a `change` listener on the file input (the moment the bytes reach the viewer)
 * and again when the viewer reports `data-status="ready"` and when the first page
 * is actually in the DOM. That removes this process's IPC and polling latency
 * from the number, so it is the time the user waits and not the time a test
 * harness waits.
 *
 * Run from the repository root:
 *   npm run bench
 *   BENCH_PROFILE=1 npm run bench   # also prints the CPU hotspots for one load
 */

import { test, type Page } from "@playwright/test";
import { unzipSync } from "fflate";
import { bigDeck, bigDocx, bigPdf } from "./fixtures";

interface ProfileNode {
  id: number;
  parent?: number;
  callFrame: { functionName: string; url: string; lineNumber: number };
}

interface Profile {
  nodes: ProfileNode[];
  samples: number[];
  timeDeltas: number[];
}

interface Marks {
  /** the file arrived at the input */
  change?: number;
  /** the viewer reported status "ready" */
  ready?: number;
  /** the first page was in the DOM */
  paint?: number;
  status?: string;
}

/**
 * Stamp the clock in the page so no test-harness latency is measured.
 *
 * The viewer is only mounted once a file has been chosen (the demo shows its drop
 * area until then), so the observer watches the document rather than waiting for
 * `.radv-root` to exist first.
 */
async function arm(page: Page): Promise<void> {
  // The demo's input is visually hidden behind its own button, so it is "attached"
  // rather than "visible".
  await page.waitForSelector('input[type="file"]', { state: "attached" });
  await page.evaluate(() => {
    const marks: Marks = {};
    (window as unknown as { __bench: Marks }).__bench = marks;

    const check = (): void => {
      // Short-circuit first: once both marks are known this runs for every
      // mutation of the whole document, which is a lot of work to do nothing.
      if (marks.ready !== undefined && marks.paint !== undefined) return;
      if (marks.change === undefined) return;
      const root = document.querySelector(".radv-root");
      const status = root?.getAttribute("data-status") ?? undefined;
      if (status) marks.status = status;
      if (status === "ready" && marks.ready === undefined) marks.ready = performance.now();
      // The first page is only in the DOM once a slide is rendered into the canvas.
      if (marks.paint === undefined && document.querySelector(".radv-canvas-inner")) {
        marks.paint = performance.now();
      }
    };

    new MutationObserver(check).observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["data-status"],
    });

    const input = document.querySelector('input[type="file"]');
    input?.addEventListener(
      "change",
      () => {
        marks.change = performance.now();
        marks.ready = undefined;
        marks.paint = undefined;
        // A frame-by-frame safety net for state that changes without a mutation
        // the observer happens to see.
        const tick = (): void => {
          check();
          if (marks.ready === undefined || marks.paint === undefined) requestAnimationFrame(tick);
        };
        requestAnimationFrame(tick);
      },
      true,
    );
  });
}

async function load(page: Page, name: string, mimeType: string, buffer: Buffer): Promise<Marks> {
  await arm(page);
  await page.locator('input[type="file"]').setInputFiles({ name, mimeType, buffer });
  await page.waitForFunction(
    () => {
      const marks = (window as unknown as { __bench: Marks }).__bench;
      return marks.ready !== undefined || marks.status === "error";
    },
    undefined,
    { timeout: 180_000 },
  );
  // The paint mark is stamped by an observer; give it a frame to land.
  await page.waitForTimeout(50);
  return page.evaluate(() => (window as unknown as { __bench: Marks }).__bench);
}

function median(values: number[]): number {
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[middle - 1]! + sorted[middle]!) / 2 : sorted[middle]!;
}

interface Case {
  name: string;
  mimeType: string;
  buffer: Buffer;
}

/**
 * Fixture sizes, in the range where load time is actually measurable.
 *
 * A forty-slide deck and a twenty-page document load in tens of milliseconds on
 * any machine, which says nothing about the files that take seconds. These sizes
 * are the ones a heavy user brings: a long deck, a book-length document, a long
 * report. Override with BENCH_SCALE to check how the cost grows.
 */
function cases(): Case[] {
  const scale = Number(process.env.BENCH_SCALE ?? 1);
  const only = process.env.BENCH_ONLY ?? "";
  // BENCH_MEDIA=heavy puts two pictures on every slide of the deck.
  const heavyMedia = process.env.BENCH_MEDIA === "heavy";
  const slides = Math.round(120 * scale);
  const paragraphs = Math.round(6000 * scale);
  const pages = Math.round(600 * scale);
  const all: Case[] = [
    {
      name: `deck-${slides}${heavyMedia ? "-heavy" : ""}.pptx`,
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      buffer: Buffer.from(bigDeck(slides, heavyMedia)),
    },
    {
      name: `document-${paragraphs}.docx`,
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      buffer: Buffer.from(bigDocx(paragraphs)),
    },
    { name: `report-${pages}.pdf`, mimeType: "application/pdf", buffer: Buffer.from(bigPdf(pages)) },
  ];
  return only ? all.filter((item) => item.name.includes(only)) : all;
}

const RUNS = Number(process.env.BENCH_RUNS ?? 3);

test("median load time per format", async ({ page }) => {
  const fixtures = cases();
  for (const item of fixtures) {
    // Uncompressed size is what the parsers actually read, so both are reported.
    // (A PDF is not a package, so it only has the one number.)
    let raw = 0;
    try {
      raw = Object.values(unzipSync(new Uint8Array(item.buffer))).reduce(
        (sum, bytes) => sum + bytes.length,
        0,
      );
    } catch {
      raw = item.buffer.length;
    }
    console.log(
      `\n${item.name} — ${(item.buffer.length / 1024).toFixed(0)} KiB zipped, ${(raw / 1024).toFixed(0)} KiB uncompressed`,
    );
  }

  await page.goto("/");
  // Warm the dev server and the React module graph so the first run is not paying
  // for compilation.
  await load(page, "warmup.pptx", "application/octet-stream", Buffer.from([1, 2, 3]));

  for (const item of fixtures) {
    const ready: number[] = [];
    const paint: number[] = [];
    for (let run = 0; run < RUNS; run += 1) {
      await page.goto("/");
      const marks = await load(page, item.name, item.mimeType, item.buffer);
      if (marks.ready === undefined || marks.change === undefined) {
        console.log(`  !! ${item.name} did not reach ready (status=${marks.status})`);
        continue;
      }
      ready.push(marks.ready - marks.change);
      if (marks.paint !== undefined) paint.push(marks.paint - marks.change);
    }
    // How many pixels were rasterised, across the main page and its thumbnails.
    // For a PDF that is the work pdf.js does, and it is the cost the load time
    // does not show on a machine fast enough to hide it.
    const raster = await page.evaluate(() => {
      const canvases = [...document.querySelectorAll("canvas")];
      return {
        canvases: canvases.length,
        pixels: canvases.reduce((sum, canvas) => sum + canvas.width * canvas.height, 0),
        // The page count is a correctness check as much as a number: a change to
        // pagination that alters it has changed what the document looks like.
        pages: document.querySelector(".radv-counter")?.textContent?.trim() ?? "?",
      };
    });

    const readyMedian = median(ready);
    const paintMedian = paint.length > 0 ? median(paint) : Number.NaN;
    console.log(
      `BENCH ${item.name} ready=${readyMedian.toFixed(0)}ms paint=${Number.isNaN(paintMedian) ? "?" : `${paintMedian.toFixed(0)}ms`} runs=[${ready.map((v) => v.toFixed(0)).join(", ")}] raster=${raster.canvases} canvases/${(raster.pixels / 1e6).toFixed(1)} Mpx pages=${raster.pages.split("/").pop()?.trim() ?? "?"}`,
    );
  }
});

/**
 * Where the time goes, by self time.
 *
 * A CDP CPU profile is attribution rather than a stopwatch: it names the
 * functions the main thread was actually inside. Run it with BENCH_PROFILE=1 when
 * deciding what to optimise.
 */
test("cpu profile hotspots", async ({ page }) => {
  test.skip(process.env.BENCH_PROFILE !== "1", "set BENCH_PROFILE=1 to profile");

  const session = await page.context().newCDPSession(page);
  await session.send("Profiler.enable");
  await session.send("Profiler.setSamplingInterval", { interval: 200 });
  await page.goto("/");
  await load(page, "warmup.pptx", "application/octet-stream", Buffer.from([1, 2, 3]));

  for (const item of cases()) {
    await page.goto("/");
    await arm(page);
    await session.send("Profiler.start");
    await page.locator('input[type="file"]').setInputFiles({
      name: item.name,
      mimeType: item.mimeType,
      buffer: item.buffer,
    });
    await page.waitForFunction(
      () => {
        const marks = (window as unknown as { __bench: Marks }).__bench;
        return marks.ready !== undefined || marks.status === "error";
      },
      undefined,
      { timeout: 180_000 },
    );
    const { profile } = (await session.send("Profiler.stop")) as { profile: Profile };

    const byId = new Map<number, ProfileNode>(profile.nodes.map((node) => [node.id, node]));

    // A bundled anonymous frame has no name and no URL of its own; the nearest
    // named ancestor is what identifies it.
    const label = (node: ProfileNode): string => {
      const { functionName, url, lineNumber } = node.callFrame;
      if (functionName && url) return `${functionName}  ${url.replace(/^.*\//, "")}:${lineNumber + 1}`;
      if (functionName) return functionName;
      const ancestors: string[] = [];
      let parent = byId.get(node.parent ?? 0);
      while (parent && ancestors.length < 3) {
        if (parent.callFrame.functionName) ancestors.push(parent.callFrame.functionName);
        parent = byId.get(parent.parent ?? 0);
      }
      return `(anonymous) ← ${ancestors.join(" ← ") || "?"}`;
    };

    const selfTime = new Map<string, number>();
    let total = 0;
    let idle = 0;
    for (let index = 0; index < profile.samples.length; index += 1) {
      const node = byId.get(profile.samples[index]!);
      if (!node) continue;
      const ms = (profile.timeDeltas[index] ?? 0) / 1000;
      total += ms;
      const key = label(node);
      if (/\(idle\)|\(program\)|\(garbage collector\)/.test(key)) idle += ms;
      selfTime.set(key, (selfTime.get(key) ?? 0) + ms);
    }

    const top = [...selfTime.entries()]
      .filter(([key]) => !/\(idle\)|\(program\)|\(garbage collector\)/.test(key))
      .sort((a, b) => b[1] - a[1])
      .slice(0, 18);

    console.log(
      `\n=== ${item.name} — ${(total / 1000).toFixed(2)}s profiled, ${(idle / 1000).toFixed(2)}s idle — top self time ===`,
    );
    for (const [key, ms] of top) {
      console.log(`  ${ms.toFixed(0).padStart(6)}ms  ${key}`);
    }
  }
});
