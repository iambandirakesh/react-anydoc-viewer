/**
 * The samples the playground can load. Every one is generated in the browser, so
 * the demo ships no binary fixtures and never fetches a document.
 */

import { SAMPLE_FILE_NAME, buildSampleDeck } from "./sampleDeck";
import {
  DOCX_MIME,
  SAMPLE_DOCX_FILE_NAME,
  SAMPLE_PDF_FILE_NAME,
  buildSampleDocument,
  buildSamplePdf,
} from "./sampleDocument";
import { buildXlsx } from "../../../tests/fixtures/xlsx";
import { buildDashboardWorkbook } from "../../../tests/fixtures/xlsxDashboard";
import { buildTiff } from "../../../tests/fixtures/tiff";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** A workbook with the value kinds a reader meets: text, numbers, dates, percentages, formulas and merges. */
function buildSampleWorkbook(): Uint8Array {
  return buildXlsx({
    title: "Quarterly numbers",
    creator: "react-anydoc-viewer",
    sheets: [
      {
        name: "Summary",
        cols: [
          { min: 1, max: 1, width: 22 },
          { min: 2, max: 4, width: 16 },
        ],
        merges: ["A6:D6"],
        rows: [
          ["Region", "Revenue", "Growth", "Updated"],
          ["North", 128400, { type: "percent", value: 0.128 }, { type: "date", serial: 45292 }],
          ["South", 98120, { type: "percent", value: -0.041 }, { type: "date", serial: 45292 }],
          ["East", 154900, { type: "percent", value: 0.203 }, { type: "date", serial: 45292 }],
          ["Total", { type: "formula", formula: "SUM(B2:B4)", value: 381420 }, null, null],
          ["Formulas show the cached value; they are never recalculated in the browser.", null, null, null],
        ],
      },
      {
        name: "Notes",
        rows: [
          ["Read-only grid"],
          ["Dates and percentages come from the cell's number format"],
          ["Merged cells keep their spans"],
          ["Hidden sheets and charts are reported, not drawn"],
        ],
      },
    ],
  });
}

/** A two-page uncompressed RGB TIFF, so the canvas decoder has a real file to draw. */
function buildSampleTiff(): Uint8Array {
  return buildTiff({
    pages: [
      { width: 480, height: 300, rgb: [214, 69, 65] },
      { width: 480, height: 300, rgb: [38, 84, 124] },
    ],
  });
}

export interface Sample {
  id: string;
  /** visible name of the playground tab */
  label: string;
  ext: string;
  fileName: string;
  mime: string;
  /** one line under the playground about what this sample exercises */
  description: string;
  build: () => Uint8Array;
}

export const SAMPLES: readonly Sample[] = [
  {
    id: "deck",
    label: "Slides",
    ext: ".pptx",
    fileName: SAMPLE_FILE_NAME,
    mime: PPTX_MIME,
    description:
      "Text, shapes, images, a group, a table, charts, an empty slide and deliberately unsupported content. Press ▶ to present.",
    build: buildSampleDeck,
  },
  {
    id: "docx",
    label: "Document",
    ext: ".docx",
    fileName: SAMPLE_DOCX_FILE_NAME,
    mime: DOCX_MIME,
    description: "Headings, lists, a merged table, a picture and page breaks, paginated in the browser.",
    build: buildSampleDocument,
  },
  {
    id: "xlsx",
    label: "Workbook",
    ext: ".xlsx",
    fileName: "sample-workbook.xlsx",
    mime: XLSX_MIME,
    description: "Dates, percentages, formulas, merges and a second sheet. Search, select, copy or download as CSV.",
    build: buildSampleWorkbook,
  },
  {
    id: "dashboard",
    label: "Dashboard",
    ext: ".xlsx",
    fileName: "sample-dashboard.xlsx",
    mime: XLSX_MIME,
    description: "A styled workbook with fills, borders, number formats, conditional formatting and embedded charts.",
    build: buildDashboardWorkbook,
  },
  {
    id: "pdf",
    label: "PDF",
    ext: ".pdf",
    fileName: SAMPLE_PDF_FILE_NAME,
    mime: "application/pdf",
    description: "Rasterised page by page with pdf.js, which is only downloaded when a PDF is opened.",
    build: buildSamplePdf,
  },
  {
    id: "tiff",
    label: "Scan",
    ext: ".tiff",
    fileName: "sample-scan.tiff",
    mime: "image/tiff",
    description: "A multi-page TIFF, decoded lazily onto a canvas. Every page becomes a page of the viewer.",
    build: buildSampleTiff,
  },
];
