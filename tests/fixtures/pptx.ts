/**
 * Test fixture builder.
 *
 * Test presentations are generated from scratch rather than checked in, so the
 * suite ships no third-party or proprietary slides. Everything here writes real
 * OOXML: a ZIP package with `[Content_Types].xml`, package relationships, a
 * presentation part, masters, layouts, themes, slides and media.
 */

import { zipSync, strToU8 } from "fflate";

export const NS = {
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  p: "http://schemas.openxmlformats.org/presentationml/2006/main",
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  dc: "http://purl.org/dc/elements/1.1/",
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  dcterms: "http://purl.org/dc/terms/",
};

export const NS_ATTRS = `xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}" xmlns:c="${NS.c}" xmlns:mc="${NS.mc}"`;

/* --------------------------------------------------------------- media */

/** 1x1 fully-transparent PNG. */
export const TINY_PNG = base64ToBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==",
);

/** 2x2 solid-red PNG. */
export const RED_PNG = base64ToBytes(
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAEklEQVR4nGP8z8Dwn4EIwDiqEAAyMAP/Uk1RJQAAAABJRU5ErkJggg==",
);

/** Magic bytes of an EMF image: recognised, but not decodable by browsers. */
export const FAKE_EMF = (() => {
  const bytes = new Uint8Array(64);
  bytes[0] = 0x01;
  bytes[40] = 0x20;
  bytes[41] = 0x45; // " E"
  bytes[42] = 0x4d; // "M"
  bytes[43] = 0x46; // "F"
  return bytes;
})();

/** Magic bytes of a WMF placeable header. */
export const FAKE_WMF = new Uint8Array([0xd7, 0xcd, 0xc6, 0x9a, 0x00, 0x00, 0x00, 0x00]);

export function base64ToBytes(base64: string): Uint8Array {
  if (typeof atob === "function") {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }
  const buffer = (globalThis as unknown as { Buffer: { from(s: string, e: string): Uint8Array } }).Buffer;
  return new Uint8Array(buffer.from(base64, "base64"));
}

/* ---------------------------------------------------------------- theme */

export function officeThemeXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme ${NS_ATTRS} name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst>
          <a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>
          <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/></a:schemeClr></a:gs>
        </a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst>
          <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>
          <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>
        </a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

/* --------------------------------------------------------------- shapes */

export interface SpOptions {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  cx?: number;
  cy?: number;
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
  preset?: string;
  adjust?: Record<string, string>;
  fill?: string;
  line?: string;
  text?: string;
  ph?: string;
  phIdx?: number;
  txBox?: boolean;
  style?: string;
  customGeometryPath?: string;
  noTransform?: boolean;
  hidden?: boolean;
}

const EMU = 914_400;
export const inch = (value: number): number => Math.round(value * EMU);

export function sp(options: SpOptions = {}): string {
  const {
    id = 2,
    name = "Shape",
    x = inch(1),
    y = inch(1),
    cx = inch(2),
    cy = inch(1),
    rot,
    flipH,
    flipV,
    preset = "rect",
    adjust,
    fill,
    line,
    text,
    ph,
    phIdx,
    txBox,
    style,
    customGeometryPath,
    noTransform,
  } = options;

  const xfrmAttrs = [
    rot !== undefined ? ` rot="${rot}"` : "",
    flipH ? ' flipH="1"' : "",
    flipV ? ' flipV="1"' : "",
  ].join("");

  const geometry = customGeometryPath
    ? `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${customGeometryPath}</a:pathLst></a:custGeom>`
    : `<a:prstGeom prst="${preset}">${
        adjust
          ? `<a:avLst>${Object.entries(adjust)
              .map(([key, value]) => `<a:gd name="${key}" fmla="val ${value}"/>`)
              .join("")}</a:avLst>`
          : "<a:avLst/>"
      }</a:prstGeom>`;

  const xfrm = noTransform
    ? ""
    : `<a:xfrm${xfrmAttrs}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;

  return `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="${id}" name="${name}"${options.hidden ? ' hidden="1"' : ""}/>
    <p:cNvSpPr${txBox ? ' txBox="1"' : ""}/>
    <p:nvPr>${ph ? `<p:ph type="${ph}"${phIdx !== undefined ? ` idx="${phIdx}"` : ""}/>` : ""}</p:nvPr>
  </p:nvSpPr>
  <p:spPr>${xfrm}${geometry}${fill ?? ""}${line ?? ""}</p:spPr>
  ${style ?? ""}
  ${text ? `<p:txBody><a:bodyPr/><a:lstStyle/>${text}</p:txBody>` : ""}
</p:sp>`;
}

export function pic(options: {
  id?: number;
  name?: string;
  rId: string;
  x?: number;
  y?: number;
  cx?: number;
  cy?: number;
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
  crop?: { l?: number; t?: number; r?: number; b?: number };
  alt?: string;
}): string {
  const {
    id = 3,
    name = "Picture",
    rId,
    x = inch(1),
    y = inch(1),
    cx = inch(2),
    cy = inch(2),
    rot,
    flipH,
    flipV,
    crop,
    alt,
  } = options;
  const srcRect = crop
    ? `<a:srcRect${["l", "t", "r", "b"].map((k) => {
        const value = (crop as Record<string, number | undefined>)[k];
        return value === undefined ? "" : ` ${k}="${Math.round(value * 100000)}"`;
      }).join("")}/>`
    : "";
  return `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${id}" name="${name}"${alt ? ` descr="${alt}"` : ""}/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="${rId}"/>
    ${srcRect}
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm${rot !== undefined ? ` rot="${rot}"` : ""}${flipH ? ' flipH="1"' : ""}${flipV ? ' flipV="1"' : ""}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;
}

export function graphicFrame(options: {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  cx?: number;
  cy?: number;
  content: string;
}): string {
  const { id = 4, name = "Graphic Frame", x = inch(1), y = inch(1), cx = inch(4), cy = inch(2), content } = options;
  return `<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="${id}" name="${name}"/>
    <p:cNvGraphicFramePr/>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>
  ${content}
</p:graphicFrame>`;
}

export function tableFrame(
  options: { id?: number; x?: number; y?: number; cx?: number; cy?: number; table: string },
): string {
  return graphicFrame({
    id: options.id,
    x: options.x,
    y: options.y,
    cx: options.cx,
    cy: options.cy,
    name: "Table",
    content: `<a:graphic><a:graphicData uri="${NS.a.replace("/main", "/table")}">${options.table}</a:graphicData></a:graphic>`,
  });
}

export function chartFrame(options: {
  id?: number;
  rId: string;
  x?: number;
  y?: number;
  cx?: number;
  cy?: number;
}): string {
  return graphicFrame({
    id: options.id,
    x: options.x,
    y: options.y,
    cx: options.cx,
    cy: options.cy,
    name: "Chart",
    content: `<a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${options.rId}"/></a:graphicData></a:graphic>`,
  });
}

export function group(options: {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  cx?: number;
  cy?: number;
  chX?: number;
  chY?: number;
  chCx?: number;
  chCy?: number;
  children: string;
  rot?: number;
}): string {
  const {
    id = 9,
    name = "Group",
    x = 0,
    y = 0,
    cx = inch(4),
    cy = inch(2),
    chX = 0,
    chY = 0,
    chCx = inch(4),
    chCy = inch(2),
    children,
    rot,
  } = options;
  return `<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm${rot !== undefined ? ` rot="${rot}"` : ""}>
      <a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>
      <a:chOff x="${chX}" y="${chY}"/><a:chExt cx="${chCx}" cy="${chCy}"/>
    </a:xfrm>
  </p:grpSpPr>
  ${children}
</p:grpSp>`;
}

export function paragraph(
  options: {
    align?: string;
    level?: number;
    marL?: number;
    indent?: number;
    bullet?: string;
    lineSpacingPct?: number;
    spaceAfter?: number;
    runs: string;
    endParaRPr?: string;
  },
): string {
  const attrs = [
    options.level !== undefined ? ` lvl="${options.level}"` : "",
    options.align ? ` algn="${options.align}"` : "",
    options.marL !== undefined ? ` marL="${options.marL}"` : "",
    options.indent !== undefined ? ` indent="${options.indent}"` : "",
  ].join("");
  const parts: string[] = [];
  if (options.lineSpacingPct !== undefined) {
    // `spcPct` is in thousandths of a percent: 100% is 100000.
    parts.push(`<a:lnSpc><a:spcPct val="${Math.round(options.lineSpacingPct * 1000)}"/></a:lnSpc>`);
  }
  if (options.spaceAfter !== undefined) {
    parts.push(`<a:spcAft><a:spcPts val="${Math.round(options.spaceAfter * 100)}"/></a:spcAft>`);
  }
  if (options.bullet) parts.push(options.bullet);
  return `<a:p><a:pPr${attrs}>${parts.join("")}</a:pPr>${options.runs}${
    options.endParaRPr ? `<a:endParaRPr${options.endParaRPr}/>` : ""
  }</a:p>`;
}

export function run(
  options: {
    text: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: string;
    color?: string;
    schemeColor?: string;
    font?: string;
    baseline?: number;
    linkRId?: string;
    field?: string;
  },
): string {
  const attrs = [
    options.size !== undefined ? ` sz="${options.size}"` : "",
    options.bold ? ' b="1"' : "",
    options.italic ? ' i="1"' : "",
    options.underline ? ` u="${options.underline}"` : "",
    options.baseline !== undefined ? ` baseline="${options.baseline}"` : "",
  ].join("");
  const color = options.color
    ? `<a:solidFill><a:srgbClr val="${options.color}"/></a:solidFill>`
    : options.schemeColor
      ? `<a:solidFill><a:schemeClr val="${options.schemeColor}"/></a:solidFill>`
      : "";
  const font = options.font ? `<a:latin typeface="${options.font}"/>` : "";
  const link = options.linkRId ? `<a:hlinkClick r:id="${options.linkRId}"/>` : "";
  const tag = options.field ? `fld id="{${options.field}}"` : "r";
  const textTag = options.field ? "a:t" : "a:t";
  return `<a:${tag}><a:rPr lang="en-US" dirty="0"${attrs}>${link}${color}${font}</a:rPr><${textTag}>${options.text}</${textTag}></a:${tag}>`;
}

export function bulletChar(char = "\u2022"): string {
  return `<a:buFont typeface="Arial"/><a:buChar char="${char}"/>`;
}

export function bulletAutoNum(type = "arabicPeriod"): string {
  return `<a:buAutoNum type="${type}"/>`;
}

export const BULLET_NONE = "<a:buNone/>";

/* ----------------------------------------------------------------- table */

/** `cells` are complete `<a:tc>` elements produced by {@link tableCell}. */
export function tableXml(options: {
  colWidths: number[];
  rows: { height: number; cells: string[] }[];
  firstRow?: boolean;
}): string {
  return `<a:tbl>
  <a:tblPr firstRow="${options.firstRow === false ? "0" : "1"}" bandRow="1"/>
  <a:tblGrid>${options.colWidths.map((w) => `<a:gridCol w="${w}"/>`).join("")}</a:tblGrid>
  ${options.rows
    .map((row) => `<a:tr h="${row.height}">${row.cells.join("")}</a:tr>`)
    .join("")}
</a:tbl>`;
}

export function tableCell(options: {
  text?: string;
  fill?: string;
  anchor?: string;
  marL?: number;
  gridSpan?: number;
  hMerge?: boolean;
  vMerge?: boolean;
  borders?: string;
}): string {
  const attrs = [
    options.gridSpan ? ` gridSpan="${options.gridSpan}"` : "",
    options.hMerge ? ' hMerge="1"' : "",
    options.vMerge ? ' vMerge="1"' : "",
  ].join("");
  const tcPrAttrs = [
    options.marL !== undefined ? ` marL="${options.marL}"` : "",
    options.anchor ? ` anchor="${options.anchor}"` : "",
  ].join("");
  const fill = options.fill
    ? `<a:solidFill><a:srgbClr val="${options.fill}"/></a:solidFill>`
    : "";
  const tcPr = `<a:tcPr${tcPrAttrs}>${options.borders ?? ""}${fill}</a:tcPr>`;
  const txBody = `<a:txBody><a:bodyPr/><a:lstStyle/>${
    options.text ? paragraph({ runs: run({ text: options.text, size: 1200 }) }) : ""
  }</a:txBody>`;
  return `<a:tc${attrs}>${txBody}${tcPr}</a:tc>`;
}

/* ---------------------------------------------------------------- charts */

export interface ChartSeriesSpec {
  name: string;
  categories: string[];
  values: (number | "")[];
  color?: string;
}

export function chartXml(options: {
  type?: "barChart" | "lineChart" | "pieChart" | "doughnutChart" | "areaChart" | "scatterChart";
  series: ChartSeriesSpec[];
  title?: string;
  barDir?: "col" | "bar";
  showLegend?: boolean;
  showValueLabels?: boolean;
}): string {
  const type = options.type ?? "barChart";
  const seriesXml = options.series
    .map((series, index) => {
      const catCount = series.categories.length;
      const valCount = series.values.length;
      return `<c:ser>
      <c:idx val="${index}"/><c:order val="${index}"/>
      <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(series.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>
      ${series.color ? `<c:spPr><a:solidFill><a:srgbClr val="${series.color}"/></a:solidFill></c:spPr>` : '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'}
      ${
        type === "scatterChart"
          ? `<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$${catCount + 1}</c:f><c:numCache><c:ptCount val="${catCount}"/>${series.categories
              .map((value, i) => `<c:pt idx="${i}"><c:v>${Number(value)}</c:v></c:pt>`)
              .join("")}</c:numCache></c:numRef></c:xVal>
      <c:yVal><c:numRef><c:f>Sheet1!$B$2:$B$${valCount + 1}</c:f><c:numCache><c:ptCount val="${valCount}"/>${series.values
              .map((value, i) => (value === "" ? "" : `<c:pt idx="${i}"><c:v>${value}</c:v></c:pt>`))
              .join("")}</c:numCache></c:numRef></c:yVal>`
          : `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${catCount + 1}</c:f><c:strCache><c:ptCount val="${catCount}"/>${series.categories
              .map((value, i) => `<c:pt idx="${i}"><c:v>${escapeXml(value)}</c:v></c:pt>`)
              .join("")}</c:strCache></c:strRef></c:cat>
      <c:val><c:numRef><c:f>Sheet1!$B$2:$B$${valCount + 1}</c:f><c:numCache><c:ptCount val="${valCount}"/>${series.values
              .map((value, i) => (value === "" ? "" : `<c:pt idx="${i}"><c:v>${value}</c:v></c:pt>`))
              .join("")}</c:numCache></c:numRef></c:val>`
      }
    </c:ser>`;
    })
    .join("");

  const groupAttrs =
    type === "barChart"
      ? `<c:barDir val="${options.barDir ?? "col"}"/><c:grouping val="clustered"/><c:varyColors val="0"/>`
      : type === "lineChart"
        ? '<c:grouping val="standard"/><c:varyColors val="0"/><c:marker val="1"/>'
        : '<c:varyColors val="1"/>';

  const labels = options.showValueLabels ? "<c:dLbls><c:showVal val=\"1\"/></c:dLbls>" : "";

  const axes =
    type === "pieChart" || type === "doughnutChart"
      ? ""
      : `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222222"/></c:catAx>
        <c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111111111"/></c:valAx>`;

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace ${NS_ATTRS}>
  <c:chart>
    ${
      options.title
        ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(
            options.title,
          )}</a:t></a:r></a:p></c:rich></c:tx><c:layout/></c:title>`
        : "<c:autoTitleDeleted val=\"1\"/>"
    }
    <c:plotArea>
      <c:layout/>
      <c:${type}>${groupAttrs}${seriesXml}${labels}${type === "barChart" ? '<c:gapWidth val="150"/>' : ""}${type === "doughnutChart" ? '<c:holeSize val="50"/>' : ""}${axes}</c:${type}>
      ${axes}
    </c:plotArea>
    ${options.showLegend === false ? "" : "<c:legend><c:legendPos val=\"b\"/><c:overlay val=\"0\"/></c:legend>"}
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>`;
}

export function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* -------------------------------------------------------------- package */

export interface SlideSpec {
  /** part name, e.g. `ppt/slides/slide1.xml`; order comes from sldIdLst */
  part: string;
  shapes: string;
  background?: string;
  /** extra relationships: rId -> { target, type } */
  rels?: Record<string, { target: string; type: string; external?: boolean }>;
  name?: string;
  show?: boolean;
  /** omit the layout relationship entirely */
  noLayout?: boolean;
}

export interface MediaSpec {
  part: string;
  bytes: Uint8Array;
}

export interface BuildOptions {
  slides: SlideSpec[];
  media?: MediaSpec[];
  /** parts to override or add verbatim */
  extraParts?: Record<string, string | Uint8Array>;
  /** part names to remove after assembly (simulates broken packages) */
  omitParts?: string[];
  /** slide order in `sldIdLst`, as indices into `slides` */
  order?: number[];
  width?: number;
  height?: number;
  /** use a custom slide master spTree */
  masterSpTree?: string;
  masterTxStyles?: string;
  masterClrMap?: string;
  layoutSpTree?: string;
  layoutAttrs?: string;
  coreTitle?: string;
  coreAuthor?: string;
  /** replaces `_rels/.rels` */
  rootRels?: string;
  charts?: Record<string, string>;
  /** replaces the master `<p:bg>` element contents */
  backgroundAtMaster?: string;
}

const DEFAULT_MASTER_SP_TREE = `
  <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr/>`;

const DEFAULT_MASTER_TX_STYLES = `<p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr algn="ctr"><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr marL="342900" indent="-342900" algn="l"><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/><a:defRPr sz="2000" kern="1200"/></a:lvl1pPr>
      <a:lvl2pPr marL="742950" indent="-285750" algn="l"><a:buFont typeface="Arial"/><a:buChar char="&#8211;"/><a:defRPr sz="1800" kern="1200"/></a:lvl2pPr>
      <a:lvl3pPr marL="1143000" indent="-228600" algn="l"><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/><a:defRPr sz="1600" kern="1200"/></a:lvl3pPr>
    </p:bodyStyle>
    <p:otherStyle>
      <a:lvl1pPr algn="l"><a:defRPr sz="1800" kern="1200"/></a:lvl1pPr>
    </p:otherStyle>
  </p:txStyles>`;

const DEFAULT_LAYOUT_SP_TREE = `
  <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr/>`;

export function buildPptx(options: BuildOptions): Uint8Array {
  const width = options.width ?? 12_192_000;
  const height = options.height ?? 6_858_000;
  const files: Record<string, Uint8Array> = {};
  const add = (name: string, content: string | Uint8Array) => {
    files[name] = typeof content === "string" ? strToU8(content) : content;
  };

  const slideCount = options.slides.length;
  const slides = options.slides;

  /* presentation part + rels */
  const order = options.order ?? slides.map((_, index) => index);
  const presentationRels: string[] = [];
  let nextRel = 1;
  const slideRelIds: string[] = [];
  for (const index of order) {
    const target = slides[index]!.part.replace(/^ppt\//, "");
    const rId = `rId${nextRel++}`;
    presentationRels.push(
      `<Relationship Id="${rId}" Type="${NS.r}/slide" Target="${target}"/>`,
    );
    slideRelIds.push(rId);
  }
  const masterRId = `rId${nextRel++}`;
  presentationRels.push(
    `<Relationship Id="${masterRId}" Type="${NS.r}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`,
  );
  const themeRId = `rId${nextRel++}`;
  presentationRels.push(
    `<Relationship Id="${themeRId}" Type="${NS.r}/theme" Target="theme/theme1.xml"/>`,
  );
  const presPropsRId = `rId${nextRel++}`;
  presentationRels.push(
    `<Relationship Id="${presPropsRId}" Type="${NS.r}/presProps" Target="presProps.xml"/>`,
  );

  add(
    "ppt/presentation.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation ${NS_ATTRS} saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="${masterRId}"/></p:sldMasterIdLst>
  <p:sldIdLst>${slideRelIds.map((rId, i) => `<p:sldId id="${256 + i}" r:id="${rId}"/>`).join("")}</p:sldIdLst>
  <p:sldSz cx="${width}" cy="${height}" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`,
  );
  add(
    "ppt/_rels/presentation.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}">${presentationRels.join("")}</Relationships>`,
  );

  add(
    "_rels/.rels",
    options.rootRels ??
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}">
      <Relationship Id="rId1" Type="${NS.r}/officeDocument" Target="ppt/presentation.xml"/>
      <Relationship Id="rId2" Type="${NS.r}/metadata/core-properties" Target="docProps/core.xml"/>
    </Relationships>`,
  );

  add("ppt/presProps.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr ${NS_ATTRS}/>`);

  /* master */
  add(
    "ppt/slideMasters/slideMaster1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster ${NS_ATTRS}>
  <p:cSld>
    <p:bg>${options.backgroundAtMaster ?? '<p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:bgPr>'}</p:bg>
    <p:spTree>${options.masterSpTree ?? DEFAULT_MASTER_SP_TREE}</p:spTree>
  </p:cSld>
  <p:clrMap ${options.masterClrMap ?? 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"'}/>
  ${options.masterTxStyles ?? DEFAULT_MASTER_TX_STYLES}
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`,
  );
  add(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}">
      <Relationship Id="rId1" Type="${NS.r}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
      <Relationship Id="rId2" Type="${NS.r}/theme" Target="../theme/theme1.xml"/>
    </Relationships>`,
  );

  /* layout */
  add(
    "ppt/slideLayouts/slideLayout1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout ${NS_ATTRS} ${options.layoutAttrs ?? 'type="title" preserve="1"'}>
  <p:cSld name="Title Slide">
    <p:spTree>${options.layoutSpTree ?? DEFAULT_LAYOUT_SP_TREE}</p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`,
  );
  add(
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}">
      <Relationship Id="rId1" Type="${NS.r}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
    </Relationships>`,
  );

  add("ppt/theme/theme1.xml", officeThemeXml());

  /* slides */
  for (const slide of slides) {
    const rels: string[] = [];
    if (!slide.noLayout) {
      rels.push(
        `<Relationship Id="rId1" Type="${NS.r}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`,
      );
    }
    for (const [rId, rel] of Object.entries(slide.rels ?? {})) {
      rels.push(
        `<Relationship Id="${rId}" Type="${rel.type}" Target="${rel.target}"${
          rel.external ? ' TargetMode="External"' : ""
        }/>`,
      );
    }
    add(
      slide.part,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld ${NS_ATTRS}${slide.show === false ? ' show="0"' : ""}>
  <p:cSld${slide.name ? ` name="${escapeXml(slide.name)}"` : ""}>
    ${slide.background ?? ""}
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${slide.shapes}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`,
    );
    const relsPart = slide.part.replace(/^(.*\/)([^/]+)$/, "$1_rels/$2.rels");
    add(
      relsPart,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}">${rels.join("")}</Relationships>`,
    );
  }

  /* media */
  for (const item of options.media ?? []) {
    add(item.part, item.bytes);
  }

  /* charts */
  for (const [part, xml] of Object.entries(options.charts ?? {})) {
    add(part, xml);
    add(
      part.replace(/^(.*\/)([^/]+)$/, "$1_rels/$2.rels"),
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${NS.rel}"/>`,
    );
  }

  /* core properties */
  add(
    "docProps/core.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}">
  ${options.coreTitle ? `<dc:title>${escapeXml(options.coreTitle)}</dc:title>` : ""}
  ${options.coreAuthor ? `<dc:creator>${escapeXml(options.coreAuthor)}</dc:creator>` : ""}
</cp:coreProperties>`,
  );

  /* content types */
  const overrides: string[] = [
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
    '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
    '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
    '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
  ];
  for (const slide of slides) {
    overrides.push(
      `<Override PartName="/${slide.part}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
    );
  }
  for (const part of Object.keys(options.charts ?? {})) {
    overrides.push(
      `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
    );
  }

  add(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${NS.ct}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="emf" ContentType="image/x-emf"/>
  <Default Extension="wmf" ContentType="image/x-wmf"/>
  <Default Extension="gif" ContentType="image/gif"/>
  ${overrides.join("")}
</Types>`,
  );

  for (const [name, content] of Object.entries(options.extraParts ?? {})) add(name, content);
  for (const name of options.omitParts ?? []) delete files[name];

  return zipSync(files, { level: 6 });
}
