/**
 * Page content -> page renderer.
 *
 * Parsers describe pages they cannot express as shapes with a
 * {@link PageContent} value; this is the one place that maps each kind to the
 * component that draws it. Adding a self-drawn format means adding a kind there
 * and a case here, and nothing in the parser layer ever imports a component.
 */

import type { PageContent } from "../parser/pageContent";
import type { PageRenderer } from "../renderer/pageRenderer";
import { createImagePageRenderer } from "./ImagePage";
import { createPdfPageRenderer } from "./PdfPage";
import { createSheetPageRenderer } from "./SheetPage";
import { createTextPageRenderer } from "./TextPage";
import { createTiffPageRenderer } from "./TiffPage";

export function createPageRenderer(content: PageContent): PageRenderer {
  switch (content.kind) {
    case "image":
      return createImagePageRenderer({ src: content.src, label: content.label });
    case "tiff":
      return createTiffPageRenderer(content.bytes);
    case "pdf":
      return createPdfPageRenderer(content.document);
    case "text":
      return createTextPageRenderer(content.pages);
    case "sheet":
      return createSheetPageRenderer(content.pages);
  }
}
