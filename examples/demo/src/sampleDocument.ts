/**
 * Sample Word document and PDF for the demo.
 *
 * Both are generated in the browser from the same fixture primitives the test
 * suite uses, so the demo shows off the new formats without shipping binaries.
 * The document deliberately contains a heading ladder, inline formatting, a
 * hyperlink, bulleted and numbered lists, a merged table, a picture, an explicit
 * page break and enough prose to paginate on its own.
 */

import {
  IMAGE_RELATIONSHIP,
  buildDocx,
  inlineDrawing,
  numberingXml,
  paragraph,
  stylesXml,
  wHyperlink,
  wParagraph,
  wRun,
  wTable,
  escapeXml,
} from "../../../tests/fixtures/docx";
import { buildPdf } from "../../../tests/fixtures/pdf";
import { RED_PNG } from "../../../tests/fixtures/pptx";

export const SAMPLE_DOCX_FILE_NAME = "sample-document.docx";
export const SAMPLE_PDF_FILE_NAME = "sample-document.pdf";
export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const NUM_CLICK = 1;
const NUM_STEP = 2;

function paragraphStyleRun(size: number, font: string, color?: string): string {
  return `<w:sz w:val="${size}"/>${color ? `<w:color w:val="${color}"/>` : ""}<w:rFonts w:ascii="${font}" w:hAnsi="${font}"/>`;
}

const PROSE = [
  "This page is rendered from WordprocessingML alone: the viewer parses word/document.xml, resolves the styles and numbering parts, converts the units from twips and EMU, and then lays the paragraphs out itself.",
  "Each paragraph is measured in the browser with the same styles the renderer uses, which is how the page breaks land in the same places a reader would expect. Paragraphs that are marked keep-with-next move down a page rather than being left alone at the bottom.",
  "Tables keep their merged cells, their grid spans, their shading and their measured row heights. Pictures keep their position: an inline picture takes its place in the text flow.",
  "Anything the viewer cannot draw — header and footer text, footnotes, equations, floating shapes with text wrapping — is reported in the diagnostics panel instead of disappearing silently. Turn on debug mode above to read the list for this document.",
];

export function buildSampleDocument(): Uint8Array {
  const title = wParagraph({ runs: wRun("A Word document, rendered in the browser", { size: 40, bold: true, color: "2F5496" }), style: "Title" });

  const intro = wParagraph({
    runs:
      wRun("Everything here comes from the ") +
      wRun("open-source", { bold: true }) +
      wRun(" viewer: no conversion service, no watermark, and ") +
      wRun("no file ever leaves your machine", { italic: true }) +
      wRun("."),
    spacing: `w:after="240"`,
  });

  const link = wParagraph({
    runs: wHyperlink("rId20", "Read the project on the web", { underline: "single", color: "0563C1" }),
  });

  const bullets = ["Titles, headings and body text", "Bulleted and numbered lists", "Tables, images and page breaks"].map(
    (text) => wParagraph({ runs: wRun(text), numPr: { numId: NUM_CLICK } }),
  );

  const steps = ["Parse the package", "Lay out every page", "Draw it with React"].map((text) =>
    wParagraph({ runs: wRun(text), numPr: { numId: NUM_STEP } }),
  );

  const table = wTable({
    columns: [3600, 2600, 2600],
    borders: true,
    rows: [
      {
        cells: [
          { xml: wParagraph({ runs: wRun("Format", { bold: true }) }), fill: "D9E2F3" },
          { xml: wParagraph({ runs: wRun("Parsed by", { bold: true }) }), fill: "D9E2F3", gridSpan: 2 },
        ],
      },
      {
        cells: [
          { xml: wParagraph({ runs: wRun(".docx") }) },
          { xml: wParagraph({ runs: wRun("this package's WordprocessingML parser") }), gridSpan: 2 },
        ],
      },
      {
        cells: [
          { xml: wParagraph({ runs: wRun(".pdf") }) },
          { xml: wParagraph({ runs: wRun("pdf.js, lazily loaded") }), gridSpan: 2 },
        ],
      },
      {
        cells: [
          { xml: wParagraph({ runs: wRun(".pptx") }) },
          { xml: wParagraph({ runs: wRun("the PresentationML parser, as before") }), gridSpan: 2 },
        ],
      },
    ],
  });

  const picture = wParagraph({
    runs: wRun("", {
      inline: inlineDrawing({ relId: "rId10", width: 1143000, height: 571500, description: "Sample image" }),
    }),
  });

  const pictureCaption = wParagraph({
    runs: wRun("An inline picture, placed in the flow with its own extent.", { italic: true, size: 18, color: "595959" }),
    align: "ctr",
    spacing: `w:after="240"`,
  });

  const prose = PROSE.map((text) => wParagraph({ runs: wRun(text), spacing: `w:after="160" w:line="276" w:lineRule="auto"` }));

  const secondSection =
    wParagraph({ runs: wRun("Appendix", { size: 32, bold: true, color: "1F3763" }), keepNext: true }) +
    wParagraph({ runs: wRun("Explicit page breaks and keep-with-next both start a page here.") }) +
    paragraph("Half-points, twips and EMU are all converted into CSS pixels before anything is drawn.");

  return buildDocx({
    body:
      title +
      intro +
      link +
      wParagraph({ runs: wRun("What is supported", { size: 32, bold: true, color: "2F5496" }), keepNext: true }) +
      bullets.join("") +
      wParagraph({ runs: wRun("How it works", { size: 32, bold: true, color: "2F5496" }), keepNext: true }) +
      steps.join("") +
      table +
      picture +
      pictureCaption +
      prose.join("") +
      wParagraph({ runs: wRun("", { pageBreak: true }) }) +
      secondSection,
    styles: stylesXml({
      defaultSize: 22,
      paragraphStyles: [
        { id: "Title", name: "Title", run: paragraphStyleRun(56, "Calibri Light", "2F5496"), para: `<w:spacing w:after="240"/>` },
        { id: "Heading1", name: "heading 1", run: paragraphStyleRun(32, "Calibri Light", "2F5496") },
        { id: "Heading2", name: "heading 2", run: paragraphStyleRun(26, "Calibri Light", "2F5496") },
        { id: "Quiet", name: "Quote", run: paragraphStyleRun(20, "Calibri", "404040"), para: `<w:ind w:left="720"/>` },
      ],
      characterStyles: [{ id: "Hyperlink", name: "Hyperlink", run: `<w:color w:val="0563C1"/><w:u w:val="single"/>` }],
    }),
    numbering: numberingXml([
      { abstractId: 0, numId: NUM_CLICK, levels: [{ ilvl: 0, numFmt: "bullet", lvlText: "\uF0B7", left: 720, hanging: 360 }] },
      { abstractId: 1, numId: NUM_STEP, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1.", left: 720, hanging: 360 }] },
    ]),
    relationships:
      IMAGE_RELATIONSHIP +
      `<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://github.com/iambandirakesh/react-anydoc-viewer" TargetMode="External"/>`,
    media: { "media/image1.png": RED_PNG },
    extraParts: {
      // Headers and footers are not rendered; their presence is reported instead.
      "word/header1.xml": `<?xml version="1.0"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${wParagraph(
        { runs: wRun("react-anydoc-viewer demo") },
      )}</w:hdr>`,
    },
    core: { title: "Sample Word document", creator: "react-anydoc-viewer" },
  });
}

export function buildSamplePdf(): Uint8Array {
  return buildPdf({
    title: "Sample PDF",
    author: "react-anydoc-viewer",
    pages: [
      "A PDF page, rasterised by pdf.js",
      "Each page is drawn only when it is visible",
      `Generated in the browser for this demo (${escapeXml("no binaries")})`,
    ],
  });
}
