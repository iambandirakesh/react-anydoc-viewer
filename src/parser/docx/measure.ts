/**
 * Paragraph height estimation for DOCX pagination.
 *
 * Real heights come from laying the paragraph out, which needs a DOM and the
 * renderer's styles; that measurer lives in `renderer/measureParagraph.ts` and is
 * injected by the viewer. A parse without it — server rendering, a unit test —
 * falls back to this estimator. It is deliberately crude, but it keeps
 * pagination monotonic and readable instead of collapsing a ten-page document
 * onto one overflowing page.
 */

import { DEFAULT_DOCX_FONT_SIZE, type DocxParagraphBlock, type MeasurerHandle } from "./model";

/** The renderer's default text size, used when the first run carries none. */
const RENDERER_DEFAULT_FONT_SIZE = 18;

/**
 * Estimate how many lines the paragraph needs, then multiply by the line height.
 * Character widths are approximated from the font size, which is why this is the
 * fallback and not the primary path.
 */
export function estimateParagraphHeight(block: DocxParagraphBlock, widthPx: number): number {
  const body = block.text;
  const paragraph = body.paragraphs[0];
  if (!paragraph) return 0;
  const baseSize = (body.paragraphs[0]?.runs[0]?.props.size ?? RENDERER_DEFAULT_FONT_SIZE) || DEFAULT_DOCX_FONT_SIZE;

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

/** The measurer used when no layout engine was injected. */
export function createEstimateMeasurer(): MeasurerHandle {
  return { measure: estimateParagraphHeight, dispose: () => undefined, kind: "estimate" };
}
