/**
 * Plain-text page rendering.
 *
 * A text page is drawn exactly as it is measured: monospace, `white-space: pre`,
 * no re-wrapping. Every line was already broken to the page's column count by
 * `paginate.ts` using the shared metrics, so nothing here may reflow — otherwise
 * an HTML/CSS change would start pushing lines off the bottom of the page.
 *
 * The content is real text in the DOM, so it can be selected and searched by the
 * browser; React escapes it, and no part of the file is ever interpreted as
 * markup.
 */

import type { ReactNode } from "react";
import type { PageRenderer } from "../renderer/pageRenderer";
import type { TextPage as TextPageData } from "../parser/plaintext/paginate";
import { BODY_FONT_PX, CONTENT_MARGIN_PX, LINE_HEIGHT_PX, MONOSPACE_STACK } from "../parser/plaintext/metrics";

export interface TextPageProps {
  lines: string[];
  width: number;
  height: number;
}

export function TextPage({ lines, width, height }: TextPageProps): ReactNode {
  return (
    <div
      className="radv-text-page"
      data-testid="rv-text-page"
      style={{
        width: `${width}px`,
        height: `${height}px`,
        padding: `${CONTENT_MARGIN_PX}px`,
        fontFamily: MONOSPACE_STACK,
        fontSize: `${BODY_FONT_PX}px`,
        lineHeight: `${LINE_HEIGHT_PX}px`,
      }}
    >
      {lines.join("\n")}
    </div>
  );
}

/** Build the page renderer the viewer calls for a loaded text document. */
export function createTextPageRenderer(pages: readonly TextPageData[]): PageRenderer {
  return ({ page, width, height }) => {
    const chunk = pages[page.index];
    if (!chunk) return null;
    return <TextPage lines={chunk.lines} width={width} height={height} />;
  };
}
