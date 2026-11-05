/**
 * What a format that draws its own pages hands to the viewer.
 *
 * A `.pptx` and a `.docx` become pages of shapes, which the slide renderer draws
 * from the model alone. A PDF page, a TIFF frame, a text page or a worksheet
 * cannot be expressed as shapes, so those parsers return one of these instead:
 * plain data describing the pages. Turning it into React is the component
 * layer's job (`components/pageRenderers.tsx`), which keeps every parser free of
 * UI imports and runnable without a DOM.
 */

import type { TextPage } from "./plaintext/paginate";
import type { SheetPageData } from "./xlsx/model";

/** Structurally typed so the parser layer never imports pdf.js types at runtime. */
export interface PdfPageLike {
  getViewport(params: { scale: number }): { width: number; height: number };
  render(params: {
    canvasContext: CanvasRenderingContext2D;
    viewport: unknown;
    canvas: HTMLCanvasElement;
  }): { promise: Promise<void>; cancel(): void };
  cleanup?(): void;
}

export interface PdfDocumentLike {
  numPages: number;
  getPage(pageNumber: number): Promise<PdfPageLike>;
}

export type PageContent =
  /** one browser-decodable image, already stored as a URL */
  | { kind: "image"; src: string; label?: string }
  /** a TIFF file, decoded lazily when a page is drawn */
  | { kind: "tiff"; bytes: Uint8Array }
  /** a pdf.js document proxy */
  | { kind: "pdf"; document: PdfDocumentLike }
  /** pre-broken monospace pages */
  | { kind: "text"; pages: readonly TextPage[] }
  /** worksheets (XLSX) or a delimited table (CSV) */
  | { kind: "sheet"; pages: readonly SheetPageData[] };
