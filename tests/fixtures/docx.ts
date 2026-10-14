/**
 * DOCX test fixture builder.
 *
 * Like the presentation fixtures, documents are generated rather than checked in:
 * every fixture is a real WordprocessingML package written with `fflate`, so the
 * suite can produce exactly the constructs it needs (page breaks, merges, damaged
 * relationships, hostile hyperlinks) without shipping anybody's document.
 */

import { zipSync, strToU8 } from "fflate";

export const W = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  dc: "http://purl.org/dc/elements/1.1/",
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
} as const;

export const W_NS =
  `xmlns:w="${W.w}" xmlns:r="${W.r}" xmlns:wp="${W.wp}" xmlns:a="${W.a}" xmlns:pic="${W.pic}"`;

/** Letter, 1in margins: 624 x 864 CSS px of content. */
export const US_LETTER_SECTION =
  `<w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>`;

export function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/* ------------------------------------------------------------------- runs */

export interface RunOptions {
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strike?: boolean;
  /** half-points, as `w:sz` */
  size?: number;
  color?: string;
  font?: string;
  highlight?: string;
  vertAlign?: "superscript" | "subscript";
  style?: string;
  caps?: boolean;
  smallCaps?: boolean;
  spacing?: number;
  /** emit a `<w:br w:type="page"/>` before the text */
  pageBreak?: boolean;
  /** emit a soft line break before the text */
  lineBreak?: boolean;
  /** raw XML placed inside the run (a drawing, for example) */
  inline?: string;
  /** wraps the run in `<w:hyperlink r:id="…"/>` */
  hyperlinkId?: string;
}

export function wRPr(options: RunOptions): string {
  const parts: string[] = [];
  if (options.style) parts.push(`<w:rStyle w:val="${options.style}"/>`);
  if (options.font) parts.push(`<w:rFonts w:ascii="${options.font}" w:hAnsi="${options.font}"/>`);
  if (options.bold) parts.push("<w:b/>");
  if (options.italic) parts.push("<w:i/>");
  if (options.caps) parts.push("<w:caps/>");
  if (options.smallCaps) parts.push("<w:smallCaps/>");
  if (options.strike) parts.push("<w:strike/>");
  if (options.underline) parts.push(`<w:u w:val="${options.underline}"/>`);
  if (options.color) parts.push(`<w:color w:val="${options.color}"/>`);
  if (options.spacing) parts.push(`<w:spacing w:val="${options.spacing}"/>`);
  if (options.highlight) parts.push(`<w:highlight w:val="${options.highlight}"/>`);
  if (options.vertAlign) parts.push(`<w:vertAlign w:val="${options.vertAlign}"/>`);
  if (options.size) parts.push(`<w:sz w:val="${options.size}"/>`);
  if (parts.length === 0) return "";
  return `<w:rPr>${parts.join("")}</w:rPr>`;
}

export function wRun(text: string, options: RunOptions = {}): string {
  const content = [
    options.pageBreak ? `<w:br w:type="page"/>` : "",
    options.lineBreak ? `<w:br/>` : "",
    options.inline ?? "",
    text.length > 0 ? `<w:t xml:space="preserve">${escapeXml(text)}</w:t>` : "",
  ].join("");
  const run = `<w:r>${wRPr(options)}${content}</w:r>`;
  return options.hyperlinkId ? `<w:hyperlink r:id="${options.hyperlinkId}">${run}</w:hyperlink>` : run;
}

export function wHyperlink(relId: string, text: string, options: RunOptions = {}): string {
  return `<w:hyperlink r:id="${relId}">${wRun(text, options)}</w:hyperlink>`;
}

/* ------------------------------------------------------------- paragraphs */

export interface ParagraphOptions {
  runs: string;
  style?: string;
  align?: string;
  indent?: string;
  spacing?: string;
  numPr?: { numId: number; ilvl?: number };
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  sectPr?: string;
  paragraphMark?: string;
  borders?: boolean;
  shading?: string;
}

export function wParagraph(options: ParagraphOptions): string {
  const pPr: string[] = [];
  if (options.style) pPr.push(`<w:pStyle w:val="${options.style}"/>`);
  if (options.numPr) {
    pPr.push(
      `<w:numPr><w:ilvl w:val="${options.numPr.ilvl ?? 0}"/><w:numId w:val="${options.numPr.numId}"/></w:numPr>`,
    );
  }
  if (options.pageBreakBefore) pPr.push("<w:pageBreakBefore/>");
  if (options.keepNext) pPr.push("<w:keepNext/>");
  if (options.spacing) pPr.push(`<w:spacing ${options.spacing}/>`);
  if (options.indent) pPr.push(`<w:ind ${options.indent}/>`);
  if (options.align) pPr.push(`<w:jc w:val="${options.align}"/>`);
  if (options.borders) pPr.push("<w:pBdr><w:top w:val=\"single\" w:sz=\"4\"/></w:pBdr>");
  if (options.shading) pPr.push(`<w:shd w:val="clear" w:fill="${options.shading}"/>`);
  if (options.sectPr) pPr.push(`<w:sectPr>${options.sectPr}</w:sectPr>`);
  if (options.paragraphMark) pPr.push(`<w:rPr>${options.paragraphMark}</w:rPr>`);
  return `<w:p>${pPr.length > 0 ? `<w:pPr>${pPr.join("")}</w:pPr>` : ""}${options.runs}</w:p>`;
}

export function paragraph(text: string, options: Omit<ParagraphOptions, "runs"> & { runs?: string } = {}): string {
  const { runs, ...rest } = options;
  return wParagraph({ ...rest, runs: wRun(text) + (runs ?? "") });
}

/* ----------------------------------------------------------------- images */

export function inlineDrawing(options: {
  relId: string;
  /** extent in EMU */
  width: number;
  height: number;
  description?: string;
}): string {
  return (
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${options.width}" cy="${options.height}"/>` +
    `<wp:docPr id="1" name="Picture 1" descr="${escapeXml(options.description ?? "")}"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${options.relId}"/></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${options.width}" cy="${options.height}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing>`
  );
}

export function anchoredDrawing(options: {
  relId: string;
  width: number;
  height: number;
  /** page offsets in EMU */
  x: number;
  y: number;
  relativeFrom?: string;
  wrapNone?: boolean;
  description?: string;
}): string {
  const relativeFrom = options.relativeFrom ?? "page";
  return (
    `<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${relativeFrom}"><wp:posOffset>${options.x}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="${relativeFrom}"><wp:posOffset>${options.y}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${options.width}" cy="${options.height}"/>` +
    `${options.wrapNone === false ? "<wp:wrapSquare wrapText=\"bothSides\"/>" : "<wp:wrapNone/>"}` +
    `<wp:docPr id="2" name="Picture 2" descr="${escapeXml(options.description ?? "")}"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="Picture 2"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${options.relId}"/></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${options.width}" cy="${options.height}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing>`
  );
}

/* ----------------------------------------------------------------- tables */

export interface CellOptions {
  xml?: string;
  /** twips */
  width?: number;
  gridSpan?: number;
  vMerge?: "restart" | "continue";
  fill?: string;
  borders?: boolean;
  margins?: { top?: number; left?: number; bottom?: number; right?: number };
  vAlign?: string;
}

export function wCell(options: CellOptions): string {
  const tcPr: string[] = [];
  if (options.width !== undefined) tcPr.push(`<w:tcW w:w="${options.width}" w:type="dxa"/>`);
  if (options.gridSpan && options.gridSpan > 1) tcPr.push(`<w:gridSpan w:val="${options.gridSpan}"/>`);
  if (options.vMerge) {
    tcPr.push(options.vMerge === "restart" ? `<w:vMerge w:val="restart"/>` : "<w:vMerge/>");
  }
  if (options.fill) tcPr.push(`<w:shd w:val="clear" w:fill="${options.fill}"/>`);
  if (options.borders) {
    tcPr.push(
      `<w:tcBorders>` +
        `<w:top w:val="single" w:sz="8" w:color="FF0000"/>` +
        `<w:left w:val="single" w:sz="8" w:color="0000FF"/>` +
        `<w:bottom w:val="single" w:sz="8" w:color="00FF00"/>` +
        `<w:right w:val="single" w:sz="8" w:color="000000"/>` +
        `</w:tcBorders>`,
    );
  }
  if (options.margins) {
    const mar = options.margins;
    tcPr.push(
      `<w:tcMar>` +
        `<w:top w:w="${mar.top ?? 0}" w:type="dxa"/>` +
        `<w:left w:w="${mar.left ?? 108}" w:type="dxa"/>` +
        `<w:bottom w:w="${mar.bottom ?? 0}" w:type="dxa"/>` +
        `<w:right w:w="${mar.right ?? 108}" w:type="dxa"/>` +
        `</w:tcMar>`,
    );
  }
  if (options.vAlign) tcPr.push(`<w:vAlign w:val="${options.vAlign}"/>`);
  if (options.xml === undefined && options.vMerge === "continue") {
    return `<w:tc>${tcPr.length > 0 ? `<w:tcPr>${tcPr.join("")}</w:tcPr>` : ""}${wParagraph({ runs: "" })}</w:tc>`;
  }
  return `<w:tc>${tcPr.length > 0 ? `<w:tcPr>${tcPr.join("")}</w:tcPr>` : ""}${options.xml ?? wParagraph({ runs: "" })}</w:tc>`;
}

export function wTable(options: {
  columns: number[];
  rows: { height?: number; heightRule?: string; cells: CellOptions[] }[];
  borders?: boolean;
  width?: number;
  alignment?: string;
  style?: string;
}): string {
  const tblPr: string[] = [];
  if (options.style) tblPr.push(`<w:tblStyle w:val="${options.style}"/>`);
  if (options.width !== undefined) tblPr.push(`<w:tblW w:w="${options.width}" w:type="dxa"/>`);
  if (options.alignment) tblPr.push(`<w:jc w:val="${options.alignment}"/>`);
  if (options.borders) {
    tblPr.push(
      `<w:tblBorders>` +
        `<w:top w:val="single" w:sz="6" w:color="999999"/>` +
        `<w:left w:val="single" w:sz="6" w:color="999999"/>` +
        `<w:bottom w:val="single" w:sz="6" w:color="999999"/>` +
        `<w:right w:val="single" w:sz="6" w:color="999999"/>` +
        `<w:insideH w:val="single" w:sz="6" w:color="999999"/>` +
        `<w:insideV w:val="single" w:sz="6" w:color="999999"/>` +
        `</w:tblBorders>`,
    );
  }
  const grid = `<w:tblGrid>${options.columns.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;
  const rows = options.rows
    .map((row) => {
      const trPr =
        row.height !== undefined
          ? `<w:trPr><w:trHeight w:val="${row.height}" w:hRule="${row.heightRule ?? "atLeast"}"/></w:trPr>`
          : "";
      return `<w:tr>${trPr}${row.cells.map(wCell).join("")}</w:tr>`;
    })
    .join("");
  return `<w:tbl><w:tblPr>${tblPr.join("")}</w:tblPr>${grid}${rows}</w:tbl>`;
}

/* ----------------------------------------------------------------- styles */

export function stylesXml(options: {
  defaultSize?: number;
  defaultFont?: string;
  paragraphStyles?: {
    id: string;
    name: string;
    basedOn?: string;
    run?: string;
    para?: string;
  }[];
  characterStyles?: { id: string; name: string; run?: string }[];
} = {}): string {
  const docDefaults =
    `<w:docDefaults><w:rPrDefault><w:rPr>` +
    `<w:rFonts w:ascii="${options.defaultFont ?? "Calibri"}" w:hAnsi="${options.defaultFont ?? "Calibri"}"/>` +
    `<w:sz w:val="${options.defaultSize ?? 22}"/>` +
    `</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>${"<w:spacing w:after=\"160\" w:line=\"276\" w:lineRule=\"auto\"/>"}</w:pPr></w:pPrDefault></w:docDefaults>`;

  const paragraphs = (options.paragraphStyles ?? []).map(
    (style) =>
      `<w:style w:type="paragraph" w:styleId="${style.id}"><w:name w:val="${style.name}"/>` +
      `${style.basedOn ? `<w:basedOn w:val="${style.basedOn}"/>` : ""}` +
      `${style.para ? `<w:pPr>${style.para}</w:pPr>` : ""}` +
      `${style.run ? `<w:rPr>${style.run}</w:rPr>` : ""}` +
      `</w:style>`,
  );
  const characters = (options.characterStyles ?? []).map(
    (style) =>
      `<w:style w:type="character" w:styleId="${style.id}"><w:name w:val="${style.name}"/>` +
      `${style.run ? `<w:rPr>${style.run}</w:rPr>` : ""}</w:style>`,
  );
  return `<w:styles ${W_NS}>${docDefaults}${paragraphs.join("")}${characters.join("")}</w:styles>`;
}

export interface NumberingLevelSpec {
  ilvl: number;
  numFmt: string;
  lvlText: string;
  start?: number;
  left?: number;
  hanging?: number;
  font?: string;
}

export interface NumberingListSpec {
  abstractId?: number;
  numId?: number;
  levels?: NumberingLevelSpec[];
}

/**
 * A `w:numbering` part. Pass one list or several — they share a single root, the
 * way Word writes them.
 */
export function numberingXml(lists: NumberingListSpec | NumberingListSpec[] = {}): string {
  const all = Array.isArray(lists) ? lists : [lists];
  const abstracts = all
    .map((list, index) => {
      const abstractId = list.abstractId ?? index;
      const levels = (
        list.levels ?? [{ ilvl: 0, numFmt: "bullet", lvlText: "\uF0B7", left: 720, hanging: 360 }]
      )
        .map(
          (level) =>
            `<w:lvl w:ilvl="${level.ilvl}"><w:start w:val="${level.start ?? 1}"/>` +
            `<w:numFmt w:val="${level.numFmt}"/>` +
            `${level.font ? `<w:rPr><w:rFonts w:ascii="${level.font}" w:hAnsi="${level.font}"/></w:rPr>` : ""}` +
            `<w:lvlText w:val="${escapeXml(level.lvlText)}"/>` +
            `<w:pPr><w:ind w:left="${level.left ?? 720}" w:hanging="${level.hanging ?? 360}"/></w:pPr>` +
            `</w:lvl>`,
        )
        .join("");
      return `<w:abstractNum w:abstractNumId="${abstractId}">${levels}</w:abstractNum>`;
    })
    .join("");
  const nums = all
    .map((list, index) => {
      const abstractId = list.abstractId ?? index;
      const numId = list.numId ?? index + 1;
      return `<w:num w:numId="${numId}"><w:abstractNumId w:val="${abstractId}"/></w:num>`;
    })
    .join("");
  return `<w:numbering ${W_NS}>${abstracts}${nums}</w:numbering>`;
}

/* --------------------------------------------------------------- package */

export interface DocxFixtureOptions {
  /** `w:body` children; the section properties are appended for you */
  body: string;
  styles?: string | null;
  numbering?: string | null;
  /** extra `<Relationship>` elements for the main document part */
  relationships?: string;
  /** media parts, keyed by their name under `word/` (e.g. `media/image1.png`) */
  media?: Record<string, Uint8Array>;
  /** parts added verbatim, keyed by full part name */
  extraParts?: Record<string, string | Uint8Array>;
  /** section properties (`w:pgSz`/`w:pgMar`); defaults to US Letter with 1in margins */
  section?: string;
  core?: { title?: string; creator?: string } | null;
  /** declare the document with a different part name */
  documentPart?: string;
}

function relationshipsPart(extra: string): string {
  return `<Relationships xmlns="${W.rel}">${extra}</Relationships>`;
}

export function buildDocx(options: DocxFixtureOptions): Uint8Array {
  const documentPart = options.documentPart ?? "word/document.xml";
  const media = options.media ?? {};
  const section = options.section ?? US_LETTER_SECTION;

  const files: Record<string, Uint8Array> = {};

  const overrides = [
    `<Override PartName="/${documentPart}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`,
  ];
  if (options.styles !== null) {
    overrides.push(
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`,
    );
  }
  if (options.numbering !== null) {
    overrides.push(
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`,
    );
  }
  overrides.push(
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
  );

  files["[Content_Types].xml"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="${W.ct}">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Default Extension="png" ContentType="image/png"/>` +
      overrides.join("") +
      `</Types>`,
  );

  files["_rels/.rels"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      relationshipsPart(
        `<Relationship Id="rId1" Type="${W.r}/officeDocument" Target="${documentPart}"/>` +
          `<Relationship Id="rId2" Type="${W.r}/metadata/core-properties" Target="docProps/core.xml"/>`,
      ),
  );

  files[documentPart] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document ${W_NS}><w:body>${options.body}<w:sectPr>${section}</w:sectPr></w:body></w:document>`,
  );

  files["word/_rels/document.xml.rels"] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + relationshipsPart(options.relationships ?? ""),
  );

  if (options.styles !== null) {
    files["word/styles.xml"] = strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + (options.styles ?? stylesXml()),
    );
  }
  if (options.numbering !== null) {
    files["word/numbering.xml"] = strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + (options.numbering ?? numberingXml()),
    );
  }

  const core = options.core === undefined ? { title: "Fixture document", creator: "Fixture author" } : options.core;
  if (core !== null) {
    files["docProps/core.xml"] = strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<cp:coreProperties xmlns:cp="${W.cp}" xmlns:dc="${W.dc}">` +
        `${core.title ? `<dc:title>${escapeXml(core.title)}</dc:title>` : ""}` +
        `${core.creator ? `<dc:creator>${escapeXml(core.creator)}</dc:creator>` : ""}` +
        `</cp:coreProperties>`,
    );
  }

  for (const [name, bytes] of Object.entries(media)) {
    files[`word/${name}`] = bytes;
  }
  for (const [name, content] of Object.entries(options.extraParts ?? {})) {
    files[name] = typeof content === "string" ? strToU8(content) : content;
  }

  return zipSync(files, { level: 0 });
}

/** A one-paragraph document, the smallest thing that is still a real `.docx`. */
export function minimalDocx(text = "Hello from a Word document"): Uint8Array {
  return buildDocx({ body: paragraph(text) });
}

export const IMAGE_RELATIONSHIP =
  `<Relationship Id="rId10" Type="${W.r}/image" Target="media/image1.png"/>`;
