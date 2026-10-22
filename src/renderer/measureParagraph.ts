/**
 * Paragraph measurement for flowing documents.
 *
 * A `.docx` has no page breaks for most of its paragraphs — Word computes them
 * from the laid-out text — so pagination needs the height of every paragraph at
 * the page's content width. There is exactly one way to get that right: lay the
 * paragraph out. This module builds the same markup with the same styles the
 * renderer will use (via `textStyles`), reads the height, and throws the probe
 * away again.
 *
 * Where no layout engine exists — server rendering, a unit test — no height can
 * be measured at all, and the estimator is used instead. It is deliberately
 * crude, but it keeps pagination monotonic and readable instead of collapsing a
 * ten-page document onto one overflowing page.
 */

import type { DocxParagraphBlock, ParagraphMeasurer } from "../parser/docx/model";
import { DEFAULT_DOCX_FONT_SIZE } from "../parser/docx/model";
import { baseSizeOf, paragraphStyleFor, runStyleFor, textContainerStyle } from "./textStyles";

export interface MeasurerHandle {
  measure: ParagraphMeasurer;
  dispose(): void;
  /** which strategy is in use, so the parser can say so in the diagnostics */
  kind: "layout" | "estimate";
}

/**
 * Paragraphs laid out per pass. A document is measured in slices this size so a
 * book-length file cannot build every probe at once, and each slice still costs a
 * single forced layout instead of one per paragraph.
 */
const MEASURE_PASS_LIMIT = 4000;

/** A one-line paragraph measured at this font size must land in this band. */
const CALIBRATION_FONT_PX = 100;
const CALIBRATION_MIN = 40;
const CALIBRATION_MAX = 400;

/**
 * Estimate how many lines the paragraph needs, then multiply by the line height.
 * Character widths are approximated from the font size, which is why this is the
 * fallback and not the primary path.
 */
export function estimateParagraphHeight(block: DocxParagraphBlock, widthPx: number): number {
  const body = block.text;
  const paragraph = body.paragraphs[0];
  if (!paragraph) return 0;
  const baseSize = baseSizeOf(body) || DEFAULT_DOCX_FONT_SIZE;

  let lineUnits = 0;
  let breaks = 0;
  let widestRun = 0;
  for (const run of paragraph.runs) {
    if (run.br) {
      breaks += 1;
      lineUnits = 0;
      continue;
    }
    const size = run.props.size ?? baseSize;
    const perCharacter = size * (run.props.bold ? 0.55 : 0.5);
    lineUnits += run.text.length * perCharacter;
    widestRun = Math.max(widestRun, perCharacter);
  }

  const available = Math.max(widthPx, widestRun, 1);
  const lines = Math.max(1, breaks + 1 + Math.max(0, Math.ceil(lineUnits / available) - 1));

  const spacing = paragraph.props.lineSpacing;
  const lineHeight =
    spacing?.type === "pct"
      ? baseSize * 1.2 * spacing.value
      : spacing?.type === "pt"
        ? spacing.value
        : baseSize * 1.35;

  return lines * lineHeight + (paragraph.props.spaceBefore ?? 0) + (paragraph.props.spaceAfter ?? 0);
}

/**
 * A measurer backed by real layout, or `null` when the environment cannot lay
 * text out (or lies about it, as a DOM stub does).
 */
export function createDomMeasurer(): MeasurerHandle | null {
  if (typeof document === "undefined" || !document.body) return null;

  const host = document.createElement("div");
  host.setAttribute("aria-hidden", "true");
  // Off-screen but laid out: `display: none` and `visibility: hidden` would both
  // give a zero-height box.
  host.style.cssText =
    "position:absolute;left:-100000px;top:0;width:1px;height:1px;overflow:visible;pointer-events:none;";
  document.body.appendChild(host);

  const buildProbe = (body: DocxParagraphBlock["text"], widthPx: number): HTMLElement => {
    const wrapper = document.createElement("div");
    Object.assign(wrapper.style, textContainerStyle(body));
    wrapper.style.width = `${widthPx}px`;
    wrapper.style.height = "auto";
    wrapper.style.padding = "0";

    const paragraph = body.paragraphs[0];
    const baseSize = baseSizeOf(body);
    const element = document.createElement("p");
    if (paragraph) {
      Object.assign(element.style, paragraphStyleFor(paragraph.props, body, baseSize));
      for (const run of paragraph.runs) {
        if (run.br) {
          element.appendChild(document.createElement("br"));
          continue;
        }
        if (run.text.length === 0) continue;
        const span = document.createElement("span");
        Object.assign(span.style, runStyleFor(run, baseSize, { interactive: false }));
        span.textContent = run.text;
        element.appendChild(span);
      }
    }
    wrapper.appendChild(element);
    return wrapper;
  };

  /** One paragraph: append, read, remove. */
  const probe = (body: DocxParagraphBlock["text"], widthPx: number): number => {
    const wrapper = buildProbe(body, widthPx);
    host.appendChild(wrapper);
    // Reading the height is what forces layout, so this costs one flush.
    const height = wrapper.getBoundingClientRect().height;
    host.removeChild(wrapper);
    return height;
  };

  /**
   * One slice of a batch, in one pass.
   *
   * The probes are built first, inserted with a single operation, and only then
   * read: the forced layout is paid once for the slice rather than once per
   * paragraph, which is where a long document's load time used to go. Heights are
   * read after the insertion and before any removal, so every wrapper in the slice
   * is laid out in the same pass and the numbers are the ones a single probe would
   * have produced.
   */
  const probeBatch = (bodies: DocxParagraphBlock["text"][], widthPx: number): number[] => {
    const heights: number[] = [];
    for (let start = 0; start < bodies.length; start += MEASURE_PASS_LIMIT) {
      const slice = bodies.slice(start, start + MEASURE_PASS_LIMIT);
      const wrappers = slice.map((body) => buildProbe(body, widthPx));
      const fragment = document.createDocumentFragment();
      for (const wrapper of wrappers) fragment.appendChild(wrapper);
      host.appendChild(fragment);
      for (const wrapper of wrappers) heights.push(wrapper.getBoundingClientRect().height);
      host.textContent = "";
    }
    return heights;
  };

  // Calibrate: a stub DOM answers every measurement with the same rectangle, so
  // a paragraph that must be one line tall gives it away.
  const calibrationBody = {
    anchor: "t" as const,
    wrap: "square" as const,
    insets: { l: 0, t: 0, r: 0, b: 0 },
    autofit: { type: "none" as const, fontScale: 1, lineSpacingReduction: 0 },
    vertical: "horz",
    hasText: true,
    paragraphs: [
      {
        props: {},
        runs: [{ text: "Hg", props: { size: CALIBRATION_FONT_PX, font: "Calibri" } }],
      },
    ],
  };
  const calibration = probe(calibrationBody, 1000);
  if (!(calibration > CALIBRATION_MIN) || calibration > CALIBRATION_MAX) {
    host.remove();
    return null;
  }

  const measure = (block: DocxParagraphBlock, widthPx: number): number => probe(block.text, widthPx);
  measure.measureMany = (blocks: DocxParagraphBlock[], widthPx: number): number[] =>
    probeBatch(
      blocks.map((block) => block.text),
      widthPx,
    );

  return {
    measure,
    dispose: () => host.remove(),
    kind: "layout",
  };
}

/** The measurer used for a document: real layout when possible, estimates otherwise. */
export function createParagraphMeasurer(): MeasurerHandle {
  return createDomMeasurer() ?? { measure: estimateParagraphHeight, dispose: () => undefined, kind: "estimate" };
}
