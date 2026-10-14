/**
 * Builds the demo presentation in the browser.
 *
 * The deck is generated rather than checked in so the repository ships no binary
 * assets, and it is assembled from the same OOXML primitives the test suite uses
 * (see `tests/fixtures/pptx.ts`). Every slide deliberately exercises a different
 * part of the viewer: text runs, shapes and geometry, images, grouping, tables,
 * charts, unsupported content and a genuinely empty slide.
 */

import { strToU8 } from "fflate";
import {
  NS,
  bulletAutoNum,
  bulletChar,
  BULLET_NONE,
  buildPptx,
  chartFrame,
  chartXml,
  group,
  graphicFrame,
  inch,
  paragraph,
  pic,
  run,
  sp,
  tableCell,
  tableFrame,
  tableXml,
  type SlideSpec,
} from "../../../tests/fixtures/pptx";

/* ------------------------------------------------------------------ media */

const LOGO_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120" width="120" height="120">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#4472C4"/>
      <stop offset="1" stop-color="#70AD47"/>
    </linearGradient>
  </defs>
  <rect width="120" height="120" rx="24" fill="url(#g)"/>
  <circle cx="60" cy="60" r="30" fill="#FFFFFF" opacity="0.9"/>
  <circle cx="60" cy="60" r="16" fill="#ED7D31"/>
  <path d="M20 96 L44 72 L68 96 Z" fill="#FFFFFF" opacity="0.35"/>
</svg>`;

const CHART_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 160 90" width="160" height="90">
  <rect width="160" height="90" fill="#0B2545"/>
  <rect x="14" y="52" width="14" height="26" fill="#4472C4"/>
  <rect x="34" y="38" width="14" height="40" fill="#70AD47"/>
  <rect x="54" y="26" width="14" height="52" fill="#ED7D31"/>
  <rect x="74" y="44" width="14" height="34" fill="#FFC000"/>
  <rect x="94" y="18" width="14" height="60" fill="#5B9BD5"/>
  <line x1="8" y1="78" x2="152" y2="78" stroke="#FFFFFF" stroke-opacity="0.5" stroke-width="2"/>
</svg>`;

/* ------------------------------------------------------------------ fill */

const gradient = (from: string, to: string, angle = 2700000): string =>
  `<a:gradFill rotWithShape="1"><a:gsLst>` +
  `<a:gs pos="0"><a:srgbClr val="${from}"/></a:gs>` +
  `<a:gs pos="100000"><a:srgbClr val="${to}"/></a:gs>` +
  `</a:gsLst><a:lin ang="${angle}" scaled="0"/></a:gradFill>`;

const shadow = (alpha = 42000): string =>
  `<a:effectLst><a:outerShdw blurRad="76200" dist="38100" dir="2700000" rotWithShape="0">` +
  `<a:srgbClr val="000000"><a:alpha val="${alpha}"/></a:srgbClr></a:outerShdw></a:effectLst>`;

const solidLine = (color: string, widthPt: number): string =>
  `<a:ln w="${Math.round(widthPt * 12700)}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln>`;

const title = (text: string, size: number, color = "FFFFFF"): string =>
  sp({
    id: 100,
    name: "Title",
    x: inch(0.6),
    y: inch(0.55),
    cx: inch(8.8),
    cy: inch(0.9),
    preset: "rect",
    fill: "<a:noFill/>",
    line: "<a:ln><a:noFill/></a:ln>",
    txBox: true,
    text: paragraph({
      align: "l",
      runs: run({ text, size, bold: true, color, font: "Calibri Light" }),
      endParaRPr: ` sz="${size}" b="1"`,
    }),
  });

/* ----------------------------------------------------------------- slides */

function titleSlide(): SlideSpec {
  return {
    part: "ppt/slides/slide1.xml",
    name: "Title",
    background: `<p:bg><p:bgPr>${gradient("0B2545", "1B4F72", 5400000)}</p:bgPr></p:bg>`,
    shapes: [
      sp({
        id: 2,
        name: "Accent",
        x: inch(0.6),
        y: inch(1.05),
        cx: inch(1.6),
        cy: inch(0.09),
        preset: "rect",
        fill: `<a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>`,
        line: "<a:ln><a:noFill/></a:ln>",
      }),
      sp({
        id: 3,
        name: "Headline",
        x: inch(0.6),
        y: inch(1.5),
        cx: inch(8.8),
        cy: inch(1.9),
        preset: "rect",
        fill: "<a:noFill/>",
        line: "<a:ln><a:noFill/></a:ln>",
        txBox: true,
        text: paragraph({
          align: "l",
          runs: run({
            text: "react-anydoc-viewer",
            size: 4400,
            bold: true,
            color: "FFFFFF",
            font: "Calibri Light",
          }),
        }),
      }),
      sp({
        id: 4,
        name: "Subtitle",
        x: inch(0.6),
        y: inch(3.35),
        cx: inch(7.6),
        cy: inch(0.8),
        preset: "rect",
        fill: "<a:noFill/>",
        line: "<a:ln><a:noFill/></a:ln>",
        txBox: true,
        text: paragraph({
          align: "l",
          runs: run({
            text: "A free, watermark-free PowerPoint viewer for React. Everything below is rendered from this .pptx file in your browser.",
            size: 1700,
            color: "D6E4F0",
          }),
          lineSpacingPct: 130,
        }),
      }),
      pic({
        id: 5,
        name: "Logo",
        rId: "rId2",
        x: inch(8.35),
        y: inch(4.05),
        cx: inch(1.15),
        cy: inch(1.15),
      }),
      sp({
        id: 6,
        name: "Footer",
        x: inch(0.6),
        y: inch(4.95),
        cx: inch(8.8),
        cy: inch(0.4),
        preset: "rect",
        fill: "<a:noFill/>",
        line: "<a:ln><a:noFill/></a:ln>",
        txBox: true,
        text: paragraph({
          align: "l",
          runs: run({
            text: "Generated in the browser · MIT licensed · no external services",
            size: 1200,
            color: "9FB6CC",
          }),
        }),
      }),
    ].join("\n"),
    rels: {
      rId2: { target: "../media/logo.svg", type: `${NS.r}/image` },
    },
  };
}

function textSlide(): SlideSpec {
  return {
    part: "ppt/slides/slide2.xml",
    name: "Text",
    shapes: [
      title("Text, runs and lists", 2600, "0B2545"),
      sp({
        id: 2,
        name: "Body",
        x: inch(0.8),
        y: inch(1.6),
        cx: inch(6.0),
        cy: inch(3.6),
        preset: "rect",
        fill: "<a:noFill/>",
        line: "<a:ln><a:noFill/></a:ln>",
        txBox: true,
        text: [
          paragraph({
            level: 0,
            bullet: bulletChar(),
            marL: inch(0.25),
            indent: -inch(0.25),
            runs: run({ text: "Level one bullet", size: 1800, bold: true }),
            endParaRPr: ' sz="1800" b="1"',
          }),
          paragraph({
            level: 1,
            bullet: bulletChar("–"),
            marL: inch(0.6),
            indent: -inch(0.25),
            runs: run({ text: "Nested level two with ", size: 1600 }),
            spaceAfter: 6,
          }),
          paragraph({
            level: 1,
            bullet: bulletChar("–"),
            marL: inch(0.6),
            indent: -inch(0.25),
            runs:
              run({ text: "bold", size: 1600, bold: true }) +
              run({ text: ", " }) +
              run({ text: "italic", size: 1600, italic: true }) +
              run({ text: ", " }) +
              run({ text: "underlined", size: 1600, underline: "sng" }) +
              run({ text: ", " }) +
              run({ text: "accent colour", size: 1600, schemeColor: "accent2" }) +
              run({ text: " and superscript", baseline: 30000 }),
          }),
          paragraph({
            level: 0,
            bullet: bulletAutoNum("arabicPeriod"),
            marL: inch(0.25),
            indent: -inch(0.25),
            runs: run({ text: "Automatically numbered item", size: 1800 }),
          }),
          paragraph({
            level: 0,
            bullet: bulletAutoNum("romanUcPeriod"),
            marL: inch(0.25),
            indent: -inch(0.25),
            runs: run({ text: "Roman numerals too", size: 1800 }),
          }),
          paragraph({
            level: 0,
            bullet: BULLET_NONE,
            align: "r",
            runs: run({ text: "Right aligned, no bullet", size: 1400, color: "808080" }),
          }),
        ].join("\n"),
      }),
      sp({
        id: 3,
        name: "Note",
        x: inch(7.05),
        y: inch(1.6),
        cx: inch(2.6),
        cy: inch(3.4),
        preset: "roundRect",
        adjust: { adj: "8000" },
        fill: `<a:solidFill><a:srgbClr val="EEF3FA"/></a:solidFill>`,
        line: solidLine("4472C4", 1),
        text: [
          paragraph({
            runs: run({ text: "Vertical anchor", size: 1400, bold: true, color: "1B4F72" }),
          }),
          paragraph({
            runs: run({
              text: "Text bodies honour anchoring and wrapping; overflow is clipped the way PowerPoint clips it.",
              size: 1300,
            }),
            lineSpacingPct: 120,
          }),
        ].join("\n"),
      }),
    ].join("\n"),
  };
}

function shapeSlide(): SlideSpec {
  const tile = (
    id: number,
    name: string,
    preset: string,
    x: number,
    y: number,
    fill: string,
    extra: { adjust?: Record<string, string>; rot?: number; line?: string } = {},
  ): string =>
    sp({
      id,
      name,
      x,
      y,
      cx: inch(1.5),
      cy: inch(1.1),
      preset,
      adjust: extra.adjust,
      rot: extra.rot,
      fill,
      line: extra.line ?? solidLine("FFFFFF", 1.25),
    });

  return {
    part: "ppt/slides/slide3.xml",
    name: "Shapes",
    shapes: [
      title("Shapes, fills and outlines", 2600, "0B2545"),
      tile(2, "Rect", "rect", inch(0.8), inch(1.7), `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`),
      tile(3, "Rounded", "roundRect", inch(2.5), inch(1.7), `<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>`, {
        adjust: { adj: "25000" },
      }),
      tile(4, "Ellipse", "ellipse", inch(4.2), inch(1.7), `<a:solidFill><a:schemeClr val="accent3"/></a:solidFill>`),
      tile(5, "Triangle", "triangle", inch(5.9), inch(1.7), `<a:solidFill><a:schemeClr val="accent4"/></a:solidFill>`),
      tile(6, "Arrow", "rightArrow", inch(7.6), inch(1.7), `<a:solidFill><a:schemeClr val="accent5"/></a:solidFill>`, {
        adjust: { adj1: "50000", adj2: "50000" },
      }),
      tile(7, "Gradient", "rect", inch(0.8), inch(3.05), gradient("4472C4", "70AD47")),
      tile(8, "Rotated", "chevron", inch(2.5), inch(3.05), `<a:solidFill><a:schemeClr val="accent6"/></a:solidFill>`, {
        rot: 900000,
      }),
      tile(9, "Star", "star5", inch(4.2), inch(3.05), `<a:solidFill><a:srgbClr val="FFC000"/></a:solidFill>`),
      tile(10, "Donut", "donut", inch(5.9), inch(3.05), `<a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill>`, {
        adjust: { adj: "25000" },
      }),
      sp({
        id: 11,
        name: "Shadowed",
        x: inch(7.6),
        y: inch(3.05),
        cx: inch(1.5),
        cy: inch(1.1),
        preset: "rect",
        fill: `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`,
        line: `<a:ln><a:noFill/></a:ln>${shadow()}`,
      }),
      sp({
        id: 12,
        name: "Speaker notes shape",
        x: inch(0.8),
        y: inch(4.4),
        cx: inch(8.3),
        cy: inch(0.7),
        preset: "roundRect",
        adjust: { adj: "20000" },
        fill: `<a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>`,
        line: solidLine("BFBFBF", 0.75),
        text: paragraph({
          align: "ctr",
          runs: run({
            text: "Preset geometry is drawn as real SVG paths, so fills, outlines and rotations stay crisp at any zoom.",
            size: 1300,
            color: "404040",
          }),
        }),
      }),
    ].join("\n"),
  };
}

function imageSlide(): SlideSpec {
  return {
    part: "ppt/slides/slide4.xml",
    name: "Images",
    shapes: [
      title("Images and grouping", 2600, "0B2545"),
      pic({
        id: 2,
        name: "Chart image",
        rId: "rId2",
        x: inch(0.8),
        y: inch(1.7),
        cx: inch(3.6),
        cy: inch(2.03),
      }),
      pic({
        id: 3,
        name: "Cropped chart image",
        rId: "rId2",
        x: inch(0.8),
        y: inch(4.0),
        cx: inch(3.6),
        cy: inch(1.1),
        crop: { t: 0.25, b: 0.25 },
      }),
      pic({
        id: 4,
        name: "Rotated logo",
        rId: "rId2",
        x: inch(4.7),
        y: inch(1.8),
        cx: inch(1.6),
        cy: inch(0.9),
        rot: 900000,
      }),
      group({
        id: 10,
        name: "Group",
        x: inch(4.7),
        y: inch(3.0),
        cx: inch(4.4),
        cy: inch(2.1),
        chX: 0,
        chY: 0,
        chCx: inch(4.4),
        chCy: inch(2.1),
        children: [
          sp({
            id: 11,
            name: "Group background",
            x: 0,
            y: 0,
            cx: inch(4.4),
            cy: inch(2.1),
            preset: "roundRect",
            adjust: { adj: "6000" },
            fill: `<a:solidFill><a:srgbClr val="0B2545"/></a:solidFill>`,
            line: "<a:ln><a:noFill/></a:ln>",
          }),
          sp({
            id: 12,
            name: "Group square",
            x: inch(0.25),
            y: inch(0.5),
            cx: inch(1.0),
            cy: inch(1.0),
            preset: "rect",
            fill: `<a:solidFill><a:schemeClr val="accent4"/></a:solidFill>`,
            line: "<a:ln><a:noFill/></a:ln>",
          }),
          sp({
            id: 13,
            name: "Group text",
            x: inch(1.5),
            y: inch(0.2),
            cx: inch(2.7),
            cy: inch(1.7),
            preset: "rect",
            fill: "<a:noFill/>",
            line: "<a:ln><a:noFill/></a:ln>",
            text: [
              paragraph({
                runs: run({ text: "Grouped shapes", size: 1900, bold: true, color: "FFFFFF" }),
              }),
              paragraph({
                runs: run({
                  text: "Child coordinates are mapped through the group's transform, including rotation.",
                  size: 1300,
                  color: "C7D6E8",
                }),
                lineSpacingPct: 120,
              }),
            ].join("\n"),
          }),
        ].join("\n"),
      }),
    ].join("\n"),
    rels: {
      rId2: { target: "../media/chart.svg", type: `${NS.r}/image` },
    },
  };
}

function tableSlide(): SlideSpec {
  const head = (text: string): string =>
    tableCell({ text, fill: "1B4F72", anchor: "ctr", marL: inch(0.08) });
  const cell = (text: string, fill?: string): string =>
    tableCell({ text, fill, anchor: "ctr", marL: inch(0.08) });

  return {
    part: "ppt/slides/slide5.xml",
    name: "Tables",
    shapes: [
      title("Tables, merges and hyperlinks", 2600, "0B2545"),
      tableFrame({
        id: 2,
        x: inch(0.8),
        y: inch(1.7),
        cx: inch(8.3),
        cy: inch(2.6),
        table: tableXml({
          colWidths: [inch(2.2), inch(1.6), inch(1.6), inch(1.6), inch(1.3)],
          rows: [
            {
              height: inch(0.45),
              cells: [head("Element"), head("Supported"), head("Partial"), head("Planned"), head("Notes")],
            },
            {
              height: inch(0.45),
              cells: [
                cell("Text runs"),
                cell("yes", "E8F1E4"),
                cell("—"),
                cell("—"),
                cell("bold, italic, underline, colour"),
              ],
            },
            {
              height: inch(0.45),
              cells: [cell("Shapes"), cell("yes", "E8F1E4"), cell("—"), cell("—"), cell("120+ presets")],
            },
            {
              height: inch(0.45),
              cells: [cell("Charts"), cell("—"), cell("yes", "FDF0DC"), cell("—"), cell("bar, line, pie, area, scatter")],
            },
            {
              height: inch(0.45),
              cells: [cell("Animation"), cell("—"), cell("—"), cell("yes", "EFEFEF"), cell("not rendered")],
            },
            {
              height: inch(0.45),
              cells: [
                tableCell({
                  text: "Merged cell spanning three columns (gridSpan=3)",
                  fill: "EEF3FA",
                  gridSpan: 3,
                  marL: inch(0.08),
                }),
                cell(""),
                cell(""),
                cell("merge", "EEF3FA"),
              ],
            },
          ],
        }),
      }),
      sp({
        id: 3,
        name: "Link",
        x: inch(0.8),
        y: inch(4.55),
        cx: inch(8.3),
        cy: inch(0.5),
        preset: "rect",
        fill: "<a:noFill/>",
        line: "<a:ln><a:noFill/></a:ln>",
        txBox: true,
        text: paragraph({
          runs:
            run({ text: "Hyperlinks are sanitised: " }) +
            run({ text: "open example.com", size: 1500, schemeColor: "hlink", underline: "sng", linkRId: "rId2" }),
        }),
      }),
    ].join("\n"),
    rels: {
      rId2: { target: "https://example.com/", type: `${NS.r}/hyperlink`, external: true },
    },
  };
}

function chartSlide(): SlideSpec {
  return {
    part: "ppt/slides/slide6.xml",
    name: "Charts",
    shapes: [
      title("Charts", 2600, "0B2545"),
      chartFrame({
        id: 2,
        rId: "rId2",
        x: inch(0.8),
        y: inch(1.7),
        cx: inch(4.0),
        cy: inch(3.5),
      }),
      chartFrame({
        id: 3,
        rId: "rId3",
        x: inch(5.1),
        y: inch(1.7),
        cx: inch(4.0),
        cy: inch(3.5),
      }),
    ].join("\n"),
    rels: {
      rId2: { target: "../charts/chart1.xml", type: `${NS.r}/chart` },
      rId3: { target: "../charts/chart2.xml", type: `${NS.r}/chart` },
    },
  };
}

function unsupportedSlide(): SlideSpec {
  return {
    part: "ppt/slides/slide7.xml",
    name: "Unsupported content",
    shapes: [
      title("Unsupported content is reported, not hidden", 2400, "0B2545"),
      graphicFrame({
        id: 2,
        name: "Embedded spreadsheet",
        x: inch(0.8),
        y: inch(1.7),
        cx: inch(4.0),
        cy: inch(1.6),
        content: `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj r:id="rId2" progId="Excel.Sheet.12" name="Embedded"/></a:graphicData></a:graphic>`,
      }),
      graphicFrame({
        id: 3,
        name: "SmartArt",
        x: inch(5.1),
        y: inch(1.7),
        cx: inch(4.0),
        cy: inch(1.6),
        content: `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId2"/></a:graphicData></a:graphic>`,
      }),
      graphicFrame({
        id: 4,
        name: "Video",
        x: inch(0.8),
        y: inch(3.5),
        cx: inch(4.0),
        cy: inch(1.6),
        content: `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/media"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId2"/></a:graphicData></a:graphic>`,
      }),
      graphicFrame({
        id: 5,
        name: "Chart with no relationship",
        x: inch(5.1),
        y: inch(3.5),
        cx: inch(4.0),
        cy: inch(1.6),
        content: `<a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" r:id="rId99"/></a:graphicData></a:graphic>`,
      }),
      sp({
        id: 6,
        name: "Unknown preset",
        x: inch(0.8),
        y: inch(5.35),
        cx: inch(2.0),
        cy: inch(0.7),
        preset: "futureShapeGeom",
        fill: `<a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill>`,
        line: solidLine("808080", 0.75),
      }),
    ].join("\n"),
  };
}

function emptySlide(): SlideSpec {
  return {
    part: "ppt/slides/slide8.xml",
    name: "Genuinely empty",
    shapes: "",
  };
}

/* --------------------------------------------------------------- assembly */

/** Build the demo deck. Call it from the browser; it returns the raw `.pptx` bytes. */
export function buildSampleDeck(): Uint8Array {
  return buildPptx({
    width: 12_192_000,
    height: 6_858_000,
    coreTitle: "react-anydoc-viewer demo deck",
    coreAuthor: "react-anydoc-viewer",
    media: [
      { part: "ppt/media/logo.svg", bytes: strToU8(LOGO_SVG) },
      { part: "ppt/media/chart.svg", bytes: strToU8(CHART_SVG) },
    ],
    charts: {
      "ppt/charts/chart1.xml": chartXml({
        type: "barChart",
        barDir: "col",
        title: "Downloads per month",
        showValueLabels: true,
        series: [
          {
            name: "Downloads",
            categories: ["Mar", "Apr", "May", "Jun", "Jul"],
            values: [1200, 1900, 2400, 3100, 4200],
            color: "4472C4",
          },
        ],
      }),
      "ppt/charts/chart2.xml": chartXml({
        type: "pieChart",
        title: "Rendering time",
        showLegend: true,
        series: [
          {
            name: "Share",
            categories: ["Parse", "Layout", "Paint", "Images"],
            values: [22, 31, 27, 20],
          },
        ],
      }),
    },
    slides: [
      titleSlide(),
      textSlide(),
      shapeSlide(),
      imageSlide(),
      tableSlide(),
      chartSlide(),
      unsupportedSlide(),
      emptySlide(),
    ],
  });
}

/** File name used when the demo hands the deck to the viewer as a `File`. */
export const SAMPLE_FILE_NAME = "react-anydoc-viewer-demo.pptx";
