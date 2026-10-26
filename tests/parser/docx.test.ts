import { describe, expect, it } from "vitest";
import { parsePackageBytes } from "../../src/parser/document";
import type { DocumentParseOptions, ParsedDocument } from "../../src/parser/document";
import { textBodyPlainText } from "../../src/parser/text";
import { emuToPx, twipsToEmu, twipsToPx } from "../../src/parser/units";
import type { PictureNode, ShapeNode, SlideModel, TableNode } from "../../src/model/types";
import {
  IMAGE_RELATIONSHIP,
  anchoredDrawing,
  buildDocx,
  inlineDrawing,
  numberingXml,
  paragraph,
  stylesXml,
  wHyperlink,
  wParagraph,
  wRun,
  wTable,
  US_LETTER_SECTION,
} from "../fixtures/docx";
import { TINY_PNG } from "../fixtures/pptx";
import { buildPdf } from "../fixtures/pdf";
import { buildXlsx } from "../fixtures/xlsx";
import type { DocxParagraphBlock, ParagraphMeasurer } from "../../src/parser/docx/model";

/** Letter, 1 inch margins. */
const PAGE_WIDTH_EMU = twipsToEmu(12240);
const PAGE_HEIGHT_EMU = twipsToEmu(15840);
const MARGIN_EMU = twipsToEmu(1440);
const CONTENT_WIDTH_EMU = PAGE_WIDTH_EMU - MARGIN_EMU * 2;
const CONTENT_HEIGHT_PX = twipsToPx(15840 - 2880); // 864

function parse(bytes: Uint8Array, options: DocumentParseOptions = {}): ParsedDocument {
  return parsePackageBytes(bytes, options);
}

/** Every block whose text a reader can see, in order. */
function pageText(slide: SlideModel): string {
  return slide.shapes
    .map((shape) => (shape.kind === "shape" && shape.text ? textBodyPlainText(shape.text) : ""))
    .filter((text) => text.length > 0)
    .join("\n");
}

function shapesOf(slide: SlideModel): ShapeNode[] {
  return slide.shapes.filter((shape): shape is ShapeNode => shape.kind === "shape");
}

/** A measurer that reports a fixed height for every paragraph. */
const uniformHeight = (height: number) => () => height;

describe("DOCX documents", () => {
  it("parses a minimal document into one page with its text", () => {
    const parsed = parse(buildDocx({ body: paragraph("Hello from a Word document") }));
    expect(parsed.format).toBe("docx");
    expect(parsed.model.slides).toHaveLength(1);
    expect(parsed.model.width).toBe(PAGE_WIDTH_EMU);
    expect(parsed.model.height).toBe(PAGE_HEIGHT_EMU);
    expect(pageText(parsed.model.slides[0] as SlideModel)).toBe("Hello from a Word document");
    // Core properties come along for the viewer's onLoad metadata.
    expect(parsed.model.title).toBe("Fixture document");
    expect(parsed.model.author).toBe("Fixture author");
    parsed.dispose();
  });

  it("reads page setup, including landscape orientation and custom margins", () => {
    const parsed = parse(
      buildDocx({
        body: paragraph("x"),
        section: `<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1134" w:bottom="720" w:left="1134"/>`,
      }),
    );
    expect(parsed.model.width).toBe(twipsToEmu(11906)); // A4 portrait
    expect(parsed.model.height).toBe(twipsToEmu(16838));
    const shape = shapesOf(parsed.model.slides[0] as SlideModel)[0] as ShapeNode;
    expect(shape.xfrm.x).toBe(twipsToEmu(1134));
    parsed.dispose();

    const landscape = parse(
      buildDocx({ body: paragraph("x"), section: `<w:pgSz w:w="11906" w:h="16838" w:orient="landscape"/>` }),
    );
    expect(landscape.model.width).toBe(twipsToEmu(16838));
    expect(landscape.model.height).toBe(twipsToEmu(11906));
    landscape.dispose();
  });

  it("keeps run formatting: bold, italic, size, colour, font, underline and highlight", () => {
    const parsed = parse(
      buildDocx({
        body: wParagraph({
          runs:
            wRun("bold", { bold: true, size: 24, color: "FF0000", font: "Arial" }) +
            wRun("italic", { italic: true, underline: "double", highlight: "yellow" }) +
            wRun("super", { vertAlign: "superscript" }) +
            wRun("small", { smallCaps: true }),
        }),
        styles: null,
      }),
    );
    const runs = (shapesOf(parsed.model.slides[0] as SlideModel)[0] as ShapeNode).text?.paragraphs[0]?.runs ?? [];
    expect(runs.map((run) => run.text)).toEqual(["bold", "italic", "super", "small"]);
    // `w:sz` is in half-points: 24 half-points is 12pt, which is 16 CSS px.
    expect(runs[0]?.props).toMatchObject({ bold: true, size: 16, font: "Arial" });
    expect(runs[0]?.props.color).toEqual({ r: 255, g: 0, b: 0, a: 1 });
    expect(runs[1]?.props.underline).toBe("dbl");
    expect(runs[1]?.props.highlight).toEqual({ r: 255, g: 255, b: 0, a: 1 });
    expect(runs[2]?.props.baseline).toBeCloseTo(0.3);
    expect(runs[3]?.props.caps).toBe("small");
    parsed.dispose();
  });

  it("applies document defaults when a run carries no formatting", () => {
    const parsed = parse(
      buildDocx({ body: paragraph("default"), styles: stylesXml({ defaultSize: 24, defaultFont: "Georgia" }) }),
    );
    const run = (shapesOf(parsed.model.slides[0] as SlideModel)[0] as ShapeNode).text?.paragraphs[0]?.runs[0];
    expect(run?.props.size).toBeCloseTo(16); // 12pt
    expect(run?.props.font).toBe("Georgia");
    const paragraphProps = shapesOf(parsed.model.slides[0] as SlideModel)[0]?.text?.paragraphs[0]?.props;
    expect(paragraphProps?.spaceAfter).toBeCloseTo(10.67, 1);
    expect(paragraphProps?.lineSpacing).toEqual({ type: "pct", value: 1.15 });
    parsed.dispose();
  });

  it("resolves paragraph styles through w:basedOn and their built-in fallbacks", () => {
    const parsed = parse(
      buildDocx({
        body: wParagraph({ runs: wRun("Heading text"), style: "Heading1" }) +
          wParagraph({ runs: wRun("Body"), style: "BodyText" }),
        styles: stylesXml({
          paragraphStyles: [
            { id: "Heading1", name: "heading 1", run: `<w:sz w:val="32"/><w:b/><w:color w:val="112233"/>` },
            { id: "BodyText", name: "Body Text", basedOn: "Heading1", run: `<w:i/>` },
          ],
        }),
      }),
    );
    const shapes = shapesOf(parsed.model.slides[0] as SlideModel);
    const heading = shapes[0]?.text?.paragraphs[0]?.runs[0]?.props;
    expect(heading?.size).toBeCloseTo(21.33, 1);
    expect(heading?.bold).toBe(true);
    expect(heading?.color).toEqual({ r: 0x11, g: 0x22, b: 0x33, a: 1 });
    // The child style inherits the parent's size but adds its own italic.
    const body = shapes[1]?.text?.paragraphs[0]?.runs[0]?.props;
    expect(body?.size).toBeCloseTo(21.33, 1);
    expect(body?.italic).toBe(true);
    parsed.dispose();
  });

  it("falls back to Word's built-in heading ladder when the file has no styles", () => {
    const parsed = parse(
      buildDocx({ body: wParagraph({ runs: wRun("Title-ish"), style: "Heading2" }), styles: null }),
    );
    const props = shapesOf(parsed.model.slides[0] as SlideModel)[0]?.text?.paragraphs[0]?.runs[0]?.props;
    expect(props?.bold).toBe(true);
    expect(props?.size).toBeCloseTo(17.33, 1); // 13pt
    parsed.dispose();
  });

  it("turns numbering into bullets and automatic numbers with their indent", () => {
    const parsed = parse(
      buildDocx({
        body:
          wParagraph({ runs: wRun("first"), numPr: { numId: 1 } }) +
          wParagraph({ runs: wRun("second"), numPr: { numId: 1 } }) +
          wParagraph({ runs: wRun("numbered"), numPr: { numId: 2 } }),
        numbering: numberingXml([
          { abstractId: 0, numId: 1, levels: [{ ilvl: 0, numFmt: "bullet", lvlText: "\uF0B7", left: 720, hanging: 360 }] },
          { abstractId: 1, numId: 2, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1.", left: 1440, hanging: 360 }] },
        ]),
        styles: null,
      }),
    );
    const shapes = shapesOf(parsed.model.slides[0] as SlideModel);
    const bullet = shapes[0]?.text?.paragraphs[0]?.props.bullet;
    expect(bullet).toEqual({ type: "char", text: { char: "•", color: undefined, font: undefined } });
    expect(shapes[0]?.text?.paragraphs[0]?.props.indent).toEqual({ marL: 48, marR: 0, indent: -24 });
    const autoNum = shapes[2]?.text?.paragraphs[0]?.props.bullet;
    expect(autoNum).toEqual({
      type: "autoNum",
      scheme: "arabicPeriod",
      startAt: 1,
      text: { char: "", color: undefined, font: undefined },
    });
    expect(shapes[2]?.text?.paragraphs[0]?.props.indent).toEqual({ marL: 96, marR: 0, indent: -24 });
    parsed.dispose();
  });

  it("paginates measured blocks, an explicit break and page-break-before", () => {
    const many = Array.from({ length: 9 }, (_, index) => paragraph(`Paragraph ${index + 1}`)).join("");
    const parsed = parse(buildDocx({ body: many }), { measureParagraph: uniformHeight(100) });
    // 864px of content / 100px blocks = 8 per page.
    expect(parsed.model.slides).toHaveLength(2);
    expect(pageText(parsed.model.slides[0] as SlideModel).split("\n")).toHaveLength(8);
    expect(pageText(parsed.model.slides[1] as SlideModel)).toBe("Paragraph 9");
    // Blocks are stacked from the top margin.
    const firstY = shapesOf(parsed.model.slides[0] as SlideModel)[0]?.xfrm.y;
    const secondY = shapesOf(parsed.model.slides[0] as SlideModel)[1]?.xfrm.y;
    expect(firstY).toBe(MARGIN_EMU);
    expect(secondY).toBe(MARGIN_EMU + 100 * 9525);
    parsed.dispose();

    const explicit = parse(buildDocx({ body: paragraph("one") + wParagraph({ runs: wRun("two", { pageBreak: true }) }) }), {
      measureParagraph: uniformHeight(50),
    });
    expect(explicit.model.slides).toHaveLength(2);
    explicit.dispose();

    const before = parse(
      buildDocx({ body: paragraph("one") + wParagraph({ runs: wRun("two"), pageBreakBefore: true }) }),
      { measureParagraph: uniformHeight(50) },
    );
    expect(before.model.slides).toHaveLength(2);
    before.dispose();
  });

  it("measures paragraphs in one batch without changing the pagination", () => {
    // A measurer that can lay out a whole batch is asked to, once, and the pages
    // it produces are the ones the per-paragraph path would have produced.
    const many = Array.from({ length: 9 }, (_, index) => paragraph(`Paragraph ${index + 1}`)).join("");
    const bytes = buildDocx({
      body: `${many}${wTable({ columns: [6000], rows: [{ cells: [{ xml: paragraph("cell") }] }] })}`,
    });

    let perBlockCalls = 0;
    const perBlock: ParagraphMeasurer = () => {
      perBlockCalls += 1;
      return 100;
    };

    let batchCalls = 0;
    let batchedBlocks: DocxParagraphBlock[] = [];
    let batchedPerBlockCalls = 0;
    const batched = (() => {
      batchedPerBlockCalls += 1;
      return 100;
    }) as ParagraphMeasurer;
    batched.measureMany = (blocks: DocxParagraphBlock[]) => {
      batchCalls += 1;
      batchedBlocks = blocks;
      return blocks.map(() => 100);
    };

    const one = parse(bytes, { measureParagraph: perBlock });
    const inOnePass = parse(bytes, { measureParagraph: batched });

    // The batch carries every paragraph, in document order, and no other kind of
    // block, and the per-block entry point was never used for them.
    expect(batchCalls).toBe(1);
    expect(batchedBlocks).toHaveLength(9);
    expect(batchedBlocks.map((block) => textBodyPlainText(block.text))).toEqual(
      Array.from({ length: 9 }, (_, index) => `Paragraph ${index + 1}`),
    );
    // Nine paragraphs and one table cell (cells are measured one at a time).
    expect(perBlockCalls).toBe(10);
    // With a batch measurer, only the cell went through the per-block path.
    expect(batchedPerBlockCalls).toBe(1);

    // Same pages, same offsets, same text.
    expect(inOnePass.model.slides).toHaveLength(one.model.slides.length);
    const offsets = (parsed: ParsedDocument) =>
      parsed.model.slides.map((slide) => slide.shapes.map((shape) => shape.xfrm.y));
    expect(offsets(inOnePass)).toEqual(offsets(one));
    expect(inOnePass.model.slides.map(pageText)).toEqual(one.model.slides.map(pageText));

    one.dispose();
    inOnePass.dispose();
  });

  it("keeps a w:keepNext paragraph with the following one", () => {
    // Seven 100px blocks fill 700px of the 864px page, so the heading would fit
    // but would be left alone at the bottom.
    const filler = Array.from({ length: 7 }, (_, index) => paragraph(`Filler ${index}`)).join("");
    const parsed = parse(
      buildDocx({ body: `${filler}${wParagraph({ runs: wRun("Heading"), keepNext: true })}${paragraph("body")}` }),
      { measureParagraph: uniformHeight(100) },
    );
    expect(parsed.model.slides).toHaveLength(2);
    expect(pageText(parsed.model.slides[0] as SlideModel).split("\n")).toHaveLength(7);
    expect(pageText(parsed.model.slides[1] as SlideModel)).toBe("Heading\nbody");
    parsed.dispose();
  });

  it("builds tables with merges, fills, borders and measured row heights", () => {
    const parsed = parse(
      buildDocx({
        body: wTable({
          columns: [3000, 3000, 3000],
          rows: [
            {
              cells: [
                { xml: wParagraph({ runs: wRun("A") }), gridSpan: 2, fill: "EEEEEE", borders: true },
                { xml: wParagraph({ runs: wRun("B") }) },
              ],
            },
            {
              cells: [
                { xml: wParagraph({ runs: wRun("C") }), vMerge: "restart" },
                { xml: wParagraph({ runs: wRun("D") }) },
                { xml: wParagraph({ runs: wRun("E") }) },
              ],
            },
            {
              cells: [
                { vMerge: "continue" },
                { xml: wParagraph({ runs: wRun("F") }) },
                { xml: wParagraph({ runs: wRun("G") }) },
              ],
            },
          ],
        }),
        styles: null,
      }),
      { measureParagraph: uniformHeight(30) },
    );
    const table = (parsed.model.slides[0] as SlideModel).shapes.find(
      (shape): shape is TableNode => shape.kind === "table",
    );
    expect(table).toBeDefined();
    expect(table?.table.colWidths).toEqual([twipsToEmu(3000), twipsToEmu(3000), twipsToEmu(3000)]);
    // The vertical merge becomes a row span on the originating cell, and the
    // continuation cell is dropped so the renderer's column walk stays aligned.
    expect(table?.table.rows[1]?.cells[0]?.rowSpan).toBe(2);
    expect(table?.table.rows[2]?.cells).toHaveLength(2);
    expect(table?.table.rows[2]?.cells.map((cell) => textBodyPlainText(cell.text))).toEqual(["F", "G"]);
    // 30px of content plus Word's default 7.2px left/right cell margins.
    expect(table ? emuToPx(table.table.rows[0]?.height ?? 0) : 0).toBeCloseTo(30, 1);
    expect(table?.table.rows[0]?.cells[0]?.gridSpan).toBe(2);
    expect(table?.table.rows[0]?.cells[0]?.fill).toEqual({ type: "solid", color: { r: 0xee, g: 0xee, b: 0xee, a: 1 } });
    expect(table?.table.rows[0]?.cells[0]?.borders?.t?.color).toEqual({ r: 255, g: 0, b: 0, a: 1 });
    expect(table?.table.rows[0]?.cells[0]?.insets).toEqual({ l: 7.2, t: 0, r: 7.2, b: 0 });
    // The table keeps the width its own grid declares, which here is narrower
    // than the page's content area.
    expect(table ? emuToPx(table.xfrm.cx) : 0).toBeCloseTo(twipsToPx(9000), 1);
    expect(emuToPx(CONTENT_WIDTH_EMU)).toBeCloseTo(624, 1);
    parsed.dispose();
  });

  it("reports a table that relies on a Word table style", () => {
    const parsed = parse(
      buildDocx({
        body: wTable({
          columns: [9000],
          style: "TableGrid",
          rows: [{ cells: [{ xml: wParagraph({ runs: wRun("cell") }) }] }],
        }),
        styles: null,
      }),
      { measureParagraph: uniformHeight(20) },
    );
    const codes = parsed.model.diagnostics.map((diagnostic) => diagnostic.code);
    expect(codes).toContain("docx.table.style");
    const table = (parsed.model.slides[0] as SlideModel).shapes.find(
      (shape): shape is TableNode => shape.kind === "table",
    );
    expect(table?.table.rows[0]?.cells[0]?.borders?.t).not.toBeNull();
    parsed.dispose();
  });

  it("places inline pictures in the flow and floating pictures at their page offset", () => {
    const parsed = parse(
      buildDocx({
        body:
          paragraph("before") +
          wParagraph({
            runs: wRun("", { inline: inlineDrawing({ relId: "rId10", width: 914400, height: 457200, description: "logo" }) }),
          }) +
          paragraph("after") +
          wParagraph({
            runs: wRun(
              "",
              {
                inline: anchoredDrawing({
                  relId: "rId10",
                  width: 457200,
                  height: 457200,
                  x: 4572000,
                  y: 914400,
                  description: "floating",
                }),
              },
            ),
          }),
        relationships: IMAGE_RELATIONSHIP,
        media: { "media/image1.png": TINY_PNG },
        styles: null,
      }),
      { measureParagraph: uniformHeight(40) },
    );
    const pictures = (parsed.model.slides[0] as SlideModel).shapes.filter(
      (shape): shape is PictureNode => shape.kind === "picture",
    );
    expect(pictures).toHaveLength(2);
    expect(pictures[0]?.mediaPath).toBe("word/media/image1.png");
    expect(emuToPx(pictures[0]?.xfrm.cx ?? 0)).toBeCloseTo(96, 1);
    expect(pictures[0]?.alt).toBe("logo");
    // The inline picture takes flow space; the floating one does not.
    expect(pictures[0]?.xfrm.y).toBe(MARGIN_EMU + 40 * 9525);
    expect(pictures[1]?.xfrm.x).toBe(4572000);
    expect(pictures[1]?.xfrm.y).toBe(914400);
    expect(parsed.media.url("word/media/image1.png")).not.toBeNull();
    parsed.dispose();
  });

  it("warns about pictures whose image part is missing", () => {
    const parsed = parse(
      buildDocx({
        body: wParagraph({
          runs: wRun("", { inline: inlineDrawing({ relId: "rId99", width: 100000, height: 100000 }) }),
        }),
        styles: null,
      }),
    );
    const picture = (parsed.model.slides[0] as SlideModel).shapes.find(
      (shape): shape is PictureNode => shape.kind === "picture",
    );
    expect(picture?.mediaPath).toBeNull();
    expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain(
      "docx.picture.missing-relationship",
    );
    parsed.dispose();
  });

  it("keeps safe hyperlinks and refuses unsafe or internal ones", () => {
    const parsed = parse(
      buildDocx({
        body:
          wParagraph({ runs: wHyperlink("rId20", "site") }) +
          wParagraph({ runs: wHyperlink("rId21", "click me") }) +
          wParagraph({ runs: wHyperlink("rId22", "other file") }),
        relationships:
          `<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/docs" TargetMode="External"/>` +
          `<Relationship Id="rId21" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>` +
          `<Relationship Id="rId22" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="../other.docx"/>`,
        styles: null,
      }),
    );
    const shapes = shapesOf(parsed.model.slides[0] as SlideModel);
    expect(shapes[0]?.text?.paragraphs[0]?.runs[0]?.link).toMatchObject({
      href: "https://example.com/docs",
      safe: true,
    });
    // A blocked target keeps its text but never becomes an href.
    expect(shapes[1]?.text?.paragraphs[0]?.runs[0]?.link).toMatchObject({ href: "", safe: false });
    expect(shapes[2]?.text?.paragraphs[0]?.runs[0]?.link).toBeUndefined();
    const codes = parsed.model.diagnostics.map((diagnostic) => diagnostic.code);
    expect(codes).toContain("docx.hyperlink.blocked");
    expect(codes).toContain("docx.hyperlink.internal");
    parsed.dispose();
  });

  it("renders inserted runs and skips deleted ones", () => {
    const parsed = parse(
      buildDocx({
        body:
          wParagraph({
            runs:
              `<w:ins><w:r><w:t>kept</w:t></w:r></w:ins>` +
              `<w:del><w:r><w:delText>removed</w:delText></w:r></w:del>` +
              `<w:r><w:t xml:space="preserve"> and </w:t></w:r>` +
              `<w:r><w:instrText>PAGE</w:instrText></w:r>` +
              `<w:r><w:t>3</w:t></w:r>`,
          }),
        styles: null,
      }),
    );
    expect(pageText(parsed.model.slides[0] as SlideModel)).toBe("kept and 3");
    parsed.dispose();
  });

  it("reports content it cannot render instead of dropping it silently", () => {
    const parsed = parse(
      buildDocx({
        body:
          wParagraph({
            borders: true,
            shading: "FFEEEE",
            runs:
              `<w:r><w:pict><v:shape/></w:pict></w:r>` +
              `<w:r><w:drawing><wp:inline/></w:drawing></w:r>` +
              `<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>`,
          }) +
          wTable({
            columns: [4000],
            rows: [{ cells: [{ xml: `${wParagraph({ runs: wRun("outer") })}${wTable({ columns: [2000], rows: [{ cells: [{ xml: wParagraph({ runs: wRun("inner") }) }] }] })}` }] }],
          }) +
          wParagraph({ runs: wRun("after section"), sectPr: US_LETTER_SECTION }),
        styles: null,
        extraParts: { "word/header1.xml": "<w:hdr/>", "word/footnotes.xml": "<w:footnotes/>" },
      }),
      { measureParagraph: uniformHeight(30) },
    );
    const codes = new Set(parsed.model.diagnostics.map((diagnostic) => diagnostic.code));
    expect(codes).toContain("docx.vml");
    expect(codes).toContain("docx.paragraph.borders");
    expect(codes).toContain("docx.paragraph.shading");
    expect(codes).toContain("docx.equation");
    expect(codes).toContain("docx.table.nested");
    expect(codes).toContain("docx.header-footer");
    expect(codes).toContain("docx.notes");
    expect(codes).toContain("docx.section-break");
    // The section break still produced a page break rather than losing content.
    expect(parsed.model.slides.length).toBeGreaterThanOrEqual(2);
    parsed.dispose();
  });

  it("says so when page breaks had to be estimated", () => {
    // No measurer is injected, so the DOM probe runs; the stub DOM in the test
    // environment is rejected by the calibration and estimation takes over.
    const parsed = parse(
      buildDocx({ body: Array.from({ length: 40 }, (_, index) => paragraph(`line ${index}`)).join("") }),
    );
    expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain("docx.layout.estimated");
    expect(parsed.model.slides.length).toBeGreaterThan(1);
    parsed.dispose();
  });

  it("produces one empty page for a document with no content", () => {
    const parsed = parse(buildDocx({ body: "" }));
    expect(parsed.model.slides).toHaveLength(1);
    expect(parsed.model.slides[0]?.shapes).toHaveLength(0);
    expect(parsed.model.slides[0]?.background).toEqual({
      type: "solid",
      color: { r: 255, g: 255, b: 255, a: 1 },
    });
    parsed.dispose();
  });
});

describe("DOCX failure modes", () => {
  it("refuses a package whose document part is empty", () => {
    // The part is declared in the archive but carries no XML at all.
    const bytes = buildDocx({
      body: paragraph("x"),
      styles: null,
      extraParts: { "word/document.xml": "" },
    });
    expect(() => parse(bytes)).toThrowError(/no body/);
  });

  it("rejects malformed XML with a normalised error", async () => {
    const { loadDocument } = await import("../../src/parser/document");
    const bytes = buildDocx({ body: `<!DOCTYPE foo [<!ENTITY x "y">]>${paragraph("hi")}` });
    await expect(loadDocument(bytes)).rejects.toMatchObject({ code: "MALFORMED_PRESENTATION" });
  });

  it("routes a spreadsheet to the spreadsheet parser, not the document one", async () => {
    const { loadDocument } = await import("../../src/parser/document");
    const spreadsheet = buildXlsx({
      sheets: [{ name: "Sheet1", rows: [["Name"], ["Ada"]] }],
    });
    const parsed = await loadDocument(spreadsheet);
    expect(parsed.format).toBe("xlsx");
    expect(parsed.model.slides).toHaveLength(1);
    parsed.dispose();
  });

  it("keeps the synchronous package path free of PDF work", () => {
    expect(() => parse(buildPdf({ pages: ["one"] }))).toThrowError(/asynchronously/);
  });
});

describe("document geometry helpers", () => {
  it("converts twips the way the page setup expects", () => {
    expect(twipsToPx(1440)).toBeCloseTo(96);
    expect(emuToPx(twipsToEmu(1440))).toBeCloseTo(96);
    expect(CONTENT_HEIGHT_PX).toBeCloseTo(864);
  });
});
