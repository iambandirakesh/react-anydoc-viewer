import { describe, expect, it } from "vitest";
import {
  BULLET_NONE,
  bulletChar,
  buildPptx,
  chartFrame,
  chartXml,
  group,
  inch,
  NS,
  paragraph,
  pic,
  RED_PNG,
  run,
  sp,
  tableCell,
  tableFrame,
  tableXml,
  TINY_PNG,
} from "../fixtures/pptx";
import { parsePresentationBytes } from "../../src/parser/presentation";
import { readArchive } from "../../src/parser/archive";
import { emuToPx } from "../../src/parser/units";

const imageRel = (rId: string, target: string) => ({
  target,
  type: `${NS.r}/image`,
});

describe("presentation parsing", () => {
  it("reads slide size and a single slide with text", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            name: "Title",
            ph: "title",
            text: paragraph({ runs: run({ text: "Hello PowerPoint", size: 4000 }) }),
          }),
        },
      ],
    });

    const parsed = parsePresentationBytes(bytes);
    expect(parsed.model.slides).toHaveLength(1);
    expect(parsed.model.width).toBe(12_192_000);
    expect(parsed.model.height).toBe(6_858_000);

    const slide = parsed.model.slides[0]!;
    expect(slide.partName).toBe("ppt/slides/slide1.xml");
    const title = slide.shapes.find((shape) => shape.kind === "shape");
    expect(title?.kind).toBe("shape");
    if (title?.kind === "shape") {
      expect(title.text?.paragraphs[0]?.runs[0]?.text).toBe("Hello PowerPoint");
      // 40pt title inherited from the master title style.
      expect(title.text?.paragraphs[0]?.runs[0]?.props.size).toBeCloseTo(40 * (96 / 72), 3);
      expect(title.text?.paragraphs[0]?.props.align).toBe("ctr");
    }
  });

  it("uses sldIdLst order, not file names", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slideB.xml",
          shapes: sp({ text: paragraph({ runs: run({ text: "Second" }) }) }),
        },
        {
          part: "ppt/slides/slideA.xml",
          shapes: sp({ text: paragraph({ runs: run({ text: "First" }) }) }),
        },
      ],
      order: [1, 0],
    });
    const parsed = parsePresentationBytes(bytes);
    expect(parsed.model.slides.map((slide) => slide.partName)).toEqual([
      "ppt/slides/slideA.xml",
      "ppt/slides/slideB.xml",
    ]);
  });

  it("resolves theme colours through the colour map", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            fill: '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
            text: paragraph({
              runs: run({ text: "Themed", schemeColor: "tx1" }),
            }),
          }),
        },
      ],
    });
    const parsed = parsePresentationBytes(bytes);
    const shape = parsed.model.slides[0]!.shapes[0];
    expect(shape?.kind).toBe("shape");
    if (shape?.kind === "shape") {
      expect(shape.fill).toEqual({ type: "solid", color: { r: 0x44, g: 0x72, b: 0xc4, a: 1 } });
      expect(shape.text?.paragraphs[0]?.runs[0]?.props.color).toEqual({ r: 0, g: 0, b: 0, a: 1 });
    }
  });

  it("applies colour modifiers (tint/lumMod/alpha)", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            fill:
              '<a:solidFill><a:srgbClr val="4472C4"><a:lumMod val="50000"/><a:alpha val="50000"/></a:srgbClr></a:solidFill>',
            text: paragraph({ runs: run({ text: "x" }) }),
          }),
        },
      ],
    });
    const parsed = parsePresentationBytes(bytes);
    const shape = parsed.model.slides[0]!.shapes[0];
    if (shape?.kind === "shape" && shape.fill?.type === "solid") {
      expect(shape.fill.color.a).toBeCloseTo(0.5, 5);
      expect(shape.fill.color.r).toBeLessThan(0x44);
    } else {
      throw new Error("expected a solid fill");
    }
  });

  it("resolves pictures against their relationship and exposes media", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: `${pic({ rId: "rId2", alt: "A red square" })}${sp({
            ph: "body",
            text: paragraph({ runs: run({ text: "caption" }) }),
          })}`,
          rels: { rId2: imageRel("rId2", "../media/image1.png") },
        },
      ],
      media: [{ part: "ppt/media/image1.png", bytes: RED_PNG }],
    });
    const parsed = parsePresentationBytes(bytes);
    const picture = parsed.model.slides[0]!.shapes.find((shape) => shape.kind === "picture");
    expect(picture?.kind).toBe("picture");
    if (picture?.kind === "picture") {
      expect(picture.mediaPath).toBe("ppt/media/image1.png");
      expect(picture.alt).toBe("A red square");
      expect(picture.unrenderableMedia).toBeNull();
      const url = parsed.media.url("ppt/media/image1.png");
      expect(url).toMatch(/^(blob:|data:image\/png)/);
    }
  });

  it("reports unsupported media instead of dropping the element", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: pic({ rId: "rId2" }),
          rels: { rId2: imageRel("rId2", "../media/image1.emf") },
        },
      ],
      media: [
        {
          part: "ppt/media/image1.emf",
          bytes: new Uint8Array([0x01, 0x00, 0x00, 0x00, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0x20, 0x45, 0x4d, 0x46]),
        },
      ],
    });
    const parsed = parsePresentationBytes(bytes);
    const picture = parsed.model.slides[0]!.shapes.find((shape) => shape.kind === "picture");
    if (picture?.kind === "picture") {
      expect(picture.unrenderableMedia).toBe("EMF");
    } else {
      throw new Error("expected a picture node");
    }
    expect(parsed.model.diagnostics.some((d) => d.code === "picture.unsupported-format")).toBe(true);
  });

  it("warns but keeps going when a relationship is missing", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: pic({ rId: "rId99" }),
          rels: {},
        },
      ],
    });
    const parsed = parsePresentationBytes(bytes);
    const picture = parsed.model.slides[0]!.shapes[0];
    if (picture?.kind === "picture") {
      expect(picture.mediaPath).toBeNull();
    } else {
      throw new Error("expected a picture node");
    }
    expect(parsed.model.diagnostics.some((d) => d.code === "picture.missing-relationship")).toBe(true);
  });

  it("renders inherited master and layout shapes behind slide content", () => {
    const bytes = buildPptx({
      masterSpTree: `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
        ${sp({
          id: 40,
          name: "Master decoration",
          x: 0,
          y: 0,
          cx: inch(13.333),
          cy: inch(0.4),
          fill: '<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>',
        })}`,
      slides: [{ part: "ppt/slides/slide1.xml", shapes: sp({ text: paragraph({ runs: run({ text: "on top" }) }) }) }],
    });
    const parsed = parsePresentationBytes(bytes);
    const slide = parsed.model.slides[0]!;
    expect(slide.inheritedShapes.length).toBeGreaterThan(0);
    expect(slide.shapes.length).toBeGreaterThan(0);
  });

  it("parses tables, merges and charts", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: `${tableFrame({
            cx: inch(6),
            cy: inch(2),
            table: tableXml({
              colWidths: [inch(2), inch(2), inch(2)],
              rows: [
                {
                  height: inch(0.5),
                  cells: [
                    tableCell({ text: "Name", fill: "DDDDDD" }),
                    tableCell({ text: "Value", fill: "DDDDDD" }),
                    tableCell({ text: "Notes", fill: "DDDDDD" }),
                  ],
                },
                {
                  height: inch(0.5),
                  cells: [tableCell({ text: "Alpha" }), tableCell({ text: "1" }), tableCell({ text: "first" })],
                },
                {
                  height: inch(0.5),
                  cells: [tableCell({ text: "Beta" }), tableCell({ text: "2" }), tableCell({ text: "second" })],
                },
              ],
            }),
          })}${chartFrame({ rId: "rId2", y: inch(3), cx: inch(6), cy: inch(3) })}`,
          rels: { rId2: { target: "../charts/chart1.xml", type: `${NS.r}/chart` } },
        },
      ],
      charts: {
        "ppt/charts/chart1.xml": chartXml({
          type: "barChart",
          title: "Revenue",
          series: [
            { name: "2025", categories: ["Q1", "Q2", "Q3"], values: [10, 20, 30] },
            { name: "2026", categories: ["Q1", "Q2", "Q3"], values: [15, 25, 35] },
          ],
        }),
      },
    });

    const parsed = parsePresentationBytes(bytes);
    const slide = parsed.model.slides[0]!;
    const table = slide.shapes.find((shape) => shape.kind === "table");
    expect(table?.kind).toBe("table");
    if (table?.kind === "table") {
      expect(table.table.rows).toHaveLength(3);
      expect(table.table.rows[1]!.cells[0]!.text?.paragraphs[0]?.runs[0]?.text).toBe("Alpha");
      expect(table.table.colWidths).toEqual([inch(2), inch(2), inch(2)]);
    }

    const chart = slide.shapes.find((shape) => shape.kind === "chart");
    expect(chart?.kind).toBe("chart");
    if (chart?.kind === "chart") {
      expect(chart.chart.title).toBe("Revenue");
      expect(chart.chart.series).toHaveLength(2);
      expect(chart.chart.series[0]!.values).toEqual([10, 20, 30]);
      expect(chart.chart.categories).toEqual(["Q1", "Q2", "Q3"]);
      expect(chart.chart.showLegend).toBe(true);
    }
  });

  it("reports effects it cannot draw instead of dropping them silently", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            fill: `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill><a:effectLst>
              <a:glow rad="63500"><a:srgbClr val="FF0000"/></a:glow>
              <a:reflection blurRad="6350" endA="0"/>
              <a:innerShdw blurRad="38100" dist="19050" dir="2700000"><a:srgbClr val="000000"/></a:innerShdw>
            </a:effectLst>`,
          }),
        },
      ],
    });

    const codes = parsePresentationBytes(bytes).model.diagnostics.map((d) => d.code);
    expect(codes).toContain("effect.unsupported");
    expect(codes).toContain("effect.inner-shadow-approximated");

    const shape = parsePresentationBytes(bytes).model.slides[0]!.shapes[0]!;
    if (shape.kind !== "shape") throw new Error("expected a shape");
    // The inner shadow is still approximated as a shadow rather than lost.
    expect(shape.shadow?.kind).toBe("inner");
  });

  it("only gives doughnut charts a hole", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: `${chartFrame({ id: 2, rId: "rId2", cx: inch(4), cy: inch(3) })}${chartFrame({
            id: 3,
            rId: "rId3",
            x: inch(5),
            cx: inch(4),
            cy: inch(3),
          })}`,
          rels: {
            rId2: { target: "../charts/chart1.xml", type: `${NS.r}/chart` },
            rId3: { target: "../charts/chart2.xml", type: `${NS.r}/chart` },
          },
        },
      ],
      charts: {
        "ppt/charts/chart1.xml": chartXml({
          type: "pieChart",
          series: [{ name: "Share", categories: ["A", "B"], values: [60, 40] }],
        }),
        "ppt/charts/chart2.xml": chartXml({
          type: "doughnutChart",
          series: [{ name: "Share", categories: ["A", "B"], values: [60, 40] }],
        }),
      },
    });

    const charts = parsePresentationBytes(bytes).model.slides[0]!.shapes.filter((shape) => shape.kind === "chart");
    expect(charts).toHaveLength(2);
    const [pie, doughnut] = charts;
    if (pie?.kind !== "chart" || doughnut?.kind !== "chart") throw new Error("expected two charts");
    expect(pie.chart.type).toBe("pieChart");
    expect(pie.chart.holeSize).toBe(0);
    expect(doughnut.chart.holeSize).toBe(50);
  });

  it("converts percentage line spacing to a ratio", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            text: paragraph({ lineSpacingPct: 130, runs: run({ text: "Spaced out" }) }),
          }),
        },
      ],
    });

    const shape = parsePresentationBytes(bytes).model.slides[0]!.shapes[0]!;
    if (shape.kind !== "shape") throw new Error("expected a shape");
    expect(shape.text?.paragraphs[0]?.props.lineSpacing).toEqual({ type: "pct", value: 1.3 });
  });

  it("applies group child coordinate transforms", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: group({
            x: inch(1),
            y: inch(1),
            cx: inch(2),
            cy: inch(2),
            chCx: inch(4),
            chCy: inch(4),
            children: sp({ x: inch(1), y: inch(1), cx: inch(1), cy: inch(1), text: paragraph({ runs: run({ text: "in group" }) }) }),
          }),
        },
      ],
    });
    const parsed = parsePresentationBytes(bytes);
    const node = parsed.model.slides[0]!.shapes[0];
    expect(node?.kind).toBe("group");
    if (node?.kind === "group") {
      expect(node.children).toHaveLength(1);
      expect(node.chExt.cx).toBe(inch(4));
    }
  });

  it("keeps bullets from the master body style and per-paragraph overrides", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            ph: "body",
            text: `${paragraph({ runs: run({ text: "top level" }) })}${paragraph({
              level: 1,
              runs: run({ text: "nested" }),
            })}${paragraph({ bullet: BULLET_NONE, runs: run({ text: "no bullet" }) })}`,
          }),
        },
      ],
    });
    const parsed = parsePresentationBytes(bytes);
    const shape = parsed.model.slides[0]!.shapes[0];
    if (shape?.kind !== "shape") throw new Error("expected shape");
    const paragraphs = shape.text?.paragraphs ?? [];
    expect(paragraphs[0]?.props.bullet?.type).toBe("char");
    expect(paragraphs[1]?.props.bullet?.type).toBe("char");
    expect(paragraphs[2]?.props.bullet?.type).toBe("none");
    // Body level 2 text size comes from the master bodyStyle.
    expect(paragraphs[1]?.runs[0]?.props.size).toBeCloseTo(18 * (96 / 72), 3);
  });

  it("resolves numbered lists and explicit bullets", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            ph: "body",
            text: `${paragraph({ bullet: bulletChar("-"), runs: run({ text: "dash" }) })}${paragraph({
              bullet: '<a:buAutoNum type="arabicPeriod"/>',
              runs: run({ text: "one" }),
            })}`,
          }),
        },
      ],
    });
    const parsed = parsePresentationBytes(bytes);
    const shape = parsed.model.slides[0]!.shapes[0];
    if (shape?.kind !== "shape") throw new Error("expected shape");
    expect(shape.text?.paragraphs[0]?.props.bullet).toMatchObject({ type: "char" });
    expect(shape.text?.paragraphs[1]?.props.bullet).toMatchObject({
      type: "autoNum",
      scheme: "arabicPeriod",
    });
  });

  it("parses custom geometry paths", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            customGeometryPath:
              '<a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="50"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path>',
            fill: '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>',
            text: paragraph({ runs: run({ text: "custom" }) }),
          }),
        },
      ],
    });
    const parsed = parsePresentationBytes(bytes);
    const shape = parsed.model.slides[0]!.shapes[0];
    if (shape?.kind !== "shape") throw new Error("expected shape");
    expect(shape.geometry.type).toBe("custom");
    if (shape.geometry.type === "custom") {
      expect(shape.geometry.commands.map((c) => c.c)).toEqual(["M", "L", "L", "Z"]);
    }
  });

  it("reports empty slides without failing", () => {
    const bytes = buildPptx({ slides: [{ part: "ppt/slides/slide1.xml", shapes: "" }] });
    const parsed = parsePresentationBytes(bytes);
    expect(parsed.model.slides[0]!.shapes).toHaveLength(0);
    expect(parsed.model.diagnostics.some((d) => d.code === "slide.empty")).toBe(true);
  });

  it("leaves a slide with a missing layout relationship usable", () => {
    const bytes = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          noLayout: true,
          shapes: sp({ text: paragraph({ runs: run({ text: "orphan" }) }) }),
        },
      ],
    });
    const parsed = parsePresentationBytes(bytes);
    expect(parsed.model.slides).toHaveLength(1);
    expect(parsed.model.diagnostics.some((d) => d.code === "slide.no-layout")).toBe(true);
  });

  it("exposes presentation pixel dimensions matching the 16:9 ratio", () => {
    const bytes = buildPptx({
      slides: [{ part: "ppt/slides/slide1.xml", shapes: sp({ text: paragraph({ runs: run({ text: "x" }) }) }) }],
    });
    const parsed = parsePresentationBytes(bytes);
    expect(emuToPx(parsed.model.width)).toBe(1280);
    expect(emuToPx(parsed.model.height)).toBe(720);
  });

  it("does not inflate parts the renderer never needs", () => {
    const bytes = buildPptx({
      slides: [{ part: "ppt/slides/slide1.xml", shapes: sp({ text: paragraph({ runs: run({ text: "x" }) }) }) }],
      extraParts: { "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx": TINY_PNG },
    });
    const archive = readArchive(bytes);
    expect(archive.names).toContain("ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx");
    expect(archive.entries.has("ppt/embeddings/microsoft_excel_worksheet1.xlsx")).toBe(false);
  });
});
