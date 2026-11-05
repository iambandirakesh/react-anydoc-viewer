/**
 * The stage for worksheets (XLSX, CSV).
 *
 * A workbook is not navigated like slides: the sheet scrolls inside its own
 * viewport, sheets are switched with tabs along the bottom, and zoom is a plain
 * CSS scale of the grid rather than a fit to the stage. The page itself is drawn
 * by the format's page renderer, exactly as on the slide stage.
 */

import { useEffect, useState, type ReactNode, type RefObject } from "react";
import type { DocumentModel, PageModel } from "../model/types";
import type { MediaStore } from "../parser/media";
import type { SheetPageData } from "../parser/xlsx/model";
import { emuToPx } from "../parser/units";
import type { PageRenderer } from "../renderer/pageRenderer";
import type { ViewerLabels } from "../types/public";
import { MAX_ZOOM, MIN_ZOOM, type Zoom } from "../hooks/useZoom";
import { WorksheetTools, type WorksheetPresentControl } from "./WorksheetTools";

/** width of the row-number column, which a "fit width" has to make room for */
const ROW_HEADER_WIDTH_PX = 44;

let nextWorksheetPanelId = 0;

export interface WorkbookViewProps {
  model: DocumentModel;
  media: MediaStore;
  sheet: PageModel;
  sheetIndex: number;
  /** the parsed sheets, when the format provides worksheet tools (XLSX and CSV) */
  worksheetPages: readonly SheetPageData[] | undefined;
  renderPage: PageRenderer;
  showControls: boolean;
  /**
   * A presented sheet keeps only the grid and its tabs: the tabs are the one way
   * to change sheet, because the grid owns the arrow keys.
   */
  presenting: boolean;
  /** the Present button, when the host allows it */
  present: WorksheetPresentControl | undefined;
  /** remounts the worksheet tools for every new document */
  loadId: number;
  rootRef: RefObject<HTMLDivElement | null>;
  viewportRef: (element: HTMLDivElement | null) => void;
  viewportSize: { width: number; height: number };
  zoom: Zoom;
  labels: ViewerLabels;
  goToSheet: (index: number) => void;
  onRenderIssue: (code: string, message: string) => void;
}

export function WorkbookView({
  model,
  media,
  sheet,
  sheetIndex,
  worksheetPages,
  renderPage,
  showControls,
  presenting,
  present,
  loadId,
  rootRef,
  viewportRef,
  viewportSize,
  zoom,
  labels,
  goToSheet,
  onRenderIssue,
}: WorkbookViewProps): ReactNode {
  // Assign after mount so React 17 works and server/client initial markup agrees.
  const [panelId, setPanelId] = useState<string>();
  useEffect(() => {
    setPanelId(`radv-worksheet-panel-${++nextWorksheetPanelId}`);
  }, []);

  const sheetCount = model.slides.length;
  const gridScale = zoom.explicit ?? 1;

  const fitWidth = () => {
    const data = worksheetPages?.[sheetIndex];
    if (!data) return;
    const width = Math.max(
      data.widths.reduce((a, b) => a + b, ROW_HEADER_WIDTH_PX),
      ...(data.drawings ?? []).map((node) => emuToPx(node.xfrm.x + node.xfrm.cx) + ROW_HEADER_WIDTH_PX),
    );
    zoom.setZoom(Math.max(1, viewportSize.width - 2) / Math.max(1, width));
  };

  return (
    <div className="radv-workbook">
      {showControls && !presenting && worksheetPages && (
        <WorksheetTools
          key={loadId}
          pages={worksheetPages}
          index={sheetIndex}
          navigate={goToSheet}
          root={rootRef}
          fitWidth={fitWidth}
          present={present}
        />
      )}

      <div
        className="radv-workbook-viewport"
        ref={viewportRef}
        role="tabpanel"
        id={panelId}
        aria-label={sheet.name ?? undefined}
      >
        <div style={{ transform: `scale(${gridScale})`, transformOrigin: "top left" }}>
          {renderPage({
            page: sheet,
            width: Math.max(viewportSize.width, 1) / gridScale,
            height: Math.max(viewportSize.height, 1) / gridScale,
            scale: gridScale,
            interactive: true,
            media,
            labels,
            warn: onRenderIssue,
          })}
        </div>
      </div>
      <div className="radv-workbook-footer">
        <div className="radv-sheet-tabs" role="tablist" aria-label="Worksheets">
          {model.slides.map((page, index) => (
            <button
              type="button"
              role="tab"
              key={index}
              aria-selected={sheetIndex === index}
              aria-controls={panelId}
              tabIndex={sheetIndex === index ? 0 : -1}
              onClick={() => goToSheet(index)}
              onKeyDown={(event) => {
                const next =
                  event.key === "ArrowRight"
                    ? (index + 1) % sheetCount
                    : event.key === "ArrowLeft"
                      ? (index + sheetCount - 1) % sheetCount
                      : event.key === "Home"
                        ? 0
                        : event.key === "End"
                          ? sheetCount - 1
                          : null;
                if (next === null) return;
                event.preventDefault();
                event.stopPropagation();
                goToSheet(next);
                (event.currentTarget.parentElement?.children[next] as HTMLElement | undefined)?.focus();
              }}
            >
              {page.name || `Sheet ${index + 1}`}
            </button>
          ))}
        </div>
        {showControls && !presenting && (
          <div className="radv-workbook-zoom">
            <button type="button" aria-label="Zoom out" onClick={zoom.zoomOut} disabled={zoom.scale <= MIN_ZOOM}>
              −
            </button>
            <button type="button" aria-label="Reset zoom" onClick={zoom.resetZoom}>
              {Math.round(zoom.scale * 100)}%
            </button>
            <button type="button" aria-label="Zoom in" onClick={zoom.zoomIn} disabled={zoom.scale >= MAX_ZOOM}>
              +
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
