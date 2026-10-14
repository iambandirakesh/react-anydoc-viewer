import { describe, expect, it } from "vitest";
import { cleanup, fireEvent, render } from "@testing-library/react";
import { afterEach } from "vitest";
import {
  buildPptx,
  bulletChar,
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
} from "../fixtures/pptx";
import { parsePresentationBytes } from "../../src/parser/presentation";
import { SlideView } from "../../src/renderer/SlideView";

afterEach(cleanup);

function renderSlide(options: Parameters<typeof buildPptx>[0], slideIndex = 0) {
  const parsed = parsePresentationBytes(buildPptx(options));
  const slide = parsed.model.slides[slideIndex]!;
  const view = render(
    <SlideView
      slide={slide}
      media={parsed.media}
      width={1280}
      height={720}
      scope={`test-${slideIndex}`}
      interactive
    />,
  );
  return { ...view, parsed, slide };
}

describe("slide rendering", () => {
  it("renders text with inherited master formatting", () => {
    const { container, parsed } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            ph: "title",
            text: paragraph({ runs: run({ text: "Quarterly review", size: 4000, bold: true }) }),
          }),
        },
      ],
    });
    const span = container.querySelector("[data-run]");
    expect(span?.textContent).toBe("Quarterly review");
    const style = (span as HTMLElement).style;
    // 4000 hundredths of a point === 40pt === ~53.33px at 96dpi.
    expect(Number.parseFloat(style.fontSize)).toBeCloseTo(40 * (96 / 72), 3);
    expect(style.fontWeight).toBe("700");
    // Paragraph alignment comes from the master title style.
    const paragraphElement = container.querySelector("[data-paragraph]") as HTMLElement;
    expect(paragraphElement.style.textAlign).toBe("center");
    expect(parsed.model.slides[0]!.shapes.length).toBe(1);
  });

  it("draws solid fills and preset geometry as SVG paths", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            preset: "ellipse",
            fill: '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
            line: '<a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>',
            text: paragraph({ runs: run({ text: "oval" }) }),
          }),
        },
      ],
    });
    const path = container.querySelector("path");
    expect(path?.getAttribute("fill")).toBe("#4472C4");
    expect(path?.getAttribute("stroke")).toBe("#000000");
    // 25400 EMU === 2pt === 2.667px
    expect(Number.parseFloat(path?.getAttribute("stroke-width") ?? "0")).toBeCloseTo(2 * (96 / 72), 3);
    // Ellipse geometry uses arc commands rather than a plain rectangle.
    expect(path?.getAttribute("d")).toContain("A");
  });

  it("emits a gradient definition and references it", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            fill:
              '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>',
            text: paragraph({ runs: run({ text: "gradient" }) }),
          }),
        },
      ],
    });
    const gradient = container.querySelector("linearGradient");
    expect(gradient).not.toBeNull();
    const id = gradient?.getAttribute("id") ?? "";
    expect(id.startsWith("radv-g-")).toBe(true);
    expect(container.querySelector("path")?.getAttribute("fill")).toBe(`url(#${id})`);
    expect(container.querySelectorAll("stop")).toHaveLength(2);
  });

  it("applies rotation and flips through CSS transforms", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            rot: 5_400_000,
            flipH: true,
            text: paragraph({ runs: run({ text: "turned" }) }),
          }),
        },
      ],
    });
    const wrapper = container.querySelector("[data-shape-kind='shape']") as HTMLElement;
    expect(wrapper.style.transform).toContain("rotate(90deg)");
    expect(wrapper.style.transform).toContain("scale(-1, 1)");
  });

  it("renders pictures with crop-aware sizing", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: pic({ rId: "rId2", crop: { l: 25_000 / 100_000, r: 0, t: 0, b: 0 } }),
          rels: { rId2: { target: "../media/image1.png", type: `${NS.r}/image` } },
        },
      ],
      media: [{ part: "ppt/media/image1.png", bytes: RED_PNG }],
    });
    const img = container.querySelector("img");
    expect(img).not.toBeNull();
    expect(img?.getAttribute("src") ?? "").toMatch(/^(blob:|data:image\/png)/);
    // 25% cropped from the left widens the image by 1/0.75.
    const width = Number.parseFloat((img as HTMLElement).style.width);
    expect(width).toBeCloseTo((inch(2) / 9525) / 0.75, 0);
  });

  it("shows a labelled placeholder for formats browsers cannot decode", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: pic({ rId: "rId2" }),
          rels: { rId2: { target: "../media/image1.wmf", type: `${NS.r}/image` } },
        },
      ],
      media: [
        {
          part: "ppt/media/image1.wmf",
          bytes: new Uint8Array([0xd7, 0xcd, 0xc6, 0x9a, 0, 0, 0, 0]),
        },
      ],
    });
    const placeholder = container.querySelector("[data-placeholder='image']");
    expect(placeholder?.textContent).toContain("WMF");
    expect(placeholder?.textContent).toContain("cannot be displayed");
  });

  it("renders bullets and numbered list markers", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            ph: "body",
            text: `${paragraph({ bullet: bulletChar("•"), runs: run({ text: "First" }) })}${paragraph({
              bullet: '<a:buAutoNum type="arabicPeriod"/>',
              runs: run({ text: "One" }),
            })}`,
          }),
        },
      ],
    });
    const text = container.textContent ?? "";
    expect(text).toContain("•");
    expect(text).toContain("1.");
  });

  it("renders tables with cells, fills and merged spans", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: tableFrame({
            cx: inch(6),
            cy: inch(2),
            table: tableXml({
              colWidths: [inch(3), inch(3)],
              rows: [
                { height: inch(0.5), cells: [tableCell({ text: "Header", fill: "DDDDDD", gridSpan: 2 })] },
                { height: inch(0.5), cells: [tableCell({ text: "A" }), tableCell({ text: "B" })] },
              ],
            }),
          }),
        },
      ],
    });
    const cells = container.querySelectorAll("[data-cell]");
    expect(cells).toHaveLength(3);
    const header = container.querySelector("[data-cell='0-0']") as HTMLElement;
    // The spanning cell covers both columns.
    const headerWidth = Number.parseFloat(header.style.width);
    const bodyWidth = Number.parseFloat((container.querySelector("[data-cell='1-0']") as HTMLElement).style.width);
    expect(headerWidth).toBeCloseTo(bodyWidth * 2, 0);
    expect(container.textContent).toContain("Header");
  });

  it("renders charts as accessible SVG", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: chartFrame({ rId: "rId2", cx: inch(6), cy: inch(3) }),
          rels: { rId2: { target: "../charts/chart1.xml", type: `${NS.r}/chart` } },
        },
      ],
      charts: {
        "ppt/charts/chart1.xml": chartXml({
          type: "barChart",
          title: "Sales",
          series: [{ name: "2026", categories: ["Q1", "Q2"], values: [3, 7] }],
        }),
      },
    });
    const svg = container.querySelector("svg[role='img']");
    expect(svg).not.toBeNull();
    expect(svg?.getAttribute("aria-label")).toBe("Sales");
    expect(container.querySelectorAll("rect").length).toBeGreaterThan(0);
    expect(container.textContent).toContain("Q1");
  });

  it("renders a pie chart with slices", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: chartFrame({ rId: "rId2", cx: inch(6), cy: inch(3) }),
          rels: { rId2: { target: "../charts/chart1.xml", type: `${NS.r}/chart` } },
        },
      ],
      charts: {
        "ppt/charts/chart1.xml": chartXml({
          type: "pieChart",
          series: [{ name: "Share", categories: ["A", "B", "C"], values: [50, 30, 20] }],
        }),
      },
    });
    expect(container.querySelectorAll("path").length).toBeGreaterThanOrEqual(3);
    // A pie has no inner radius: every slice is a single arc closed to the centre.
    for (const path of container.querySelectorAll("path")) {
      const d = path.getAttribute("d") ?? "";
      expect(d.split("A")).toHaveLength(2);
    }
  });

  it("gives doughnut charts an inner radius", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: chartFrame({ rId: "rId2", cx: inch(6), cy: inch(3) }),
          rels: { rId2: { target: "../charts/chart1.xml", type: `${NS.r}/chart` } },
        },
      ],
      charts: {
        "ppt/charts/chart1.xml": chartXml({
          type: "doughnutChart",
          series: [{ name: "Share", categories: ["A", "B"], values: [50, 30, 20] }],
        }),
      },
    });
    const slices = [...container.querySelectorAll("path")].filter((path) =>
      (path.getAttribute("d") ?? "").startsWith("M"),
    );
    expect(slices.length).toBeGreaterThanOrEqual(3);
    for (const slice of slices) {
      // Two arcs: outer edge and inner edge.
      expect((slice.getAttribute("d") ?? "").split("A")).toHaveLength(3);
    }
  });

  it("uses longhand paragraph margins so user-agent margins cannot shift text", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            text: [
              paragraph({ runs: run({ text: "First line" }), marL: inch(0.25), indent: -inch(0.25) }),
              paragraph({ runs: run({ text: "Second line" }), spaceAfter: 8 }),
            ].join("\n"),
          }),
        },
      ],
    });
    const paragraphs = [...container.querySelectorAll("p")];
    expect(paragraphs).toHaveLength(2);
    for (const node of paragraphs) {
      // Every margin edge is set explicitly. A `margin` shorthand plus longhands is
      // not reliable across React versions (React 19 drops the shorthand), and the
      // user-agent 1em paragraph margin would otherwise move every line.
      const style = (node as HTMLElement).style;
      expect(style.marginTop).not.toBe("");
      expect(style.marginRight).not.toBe("");
      expect(style.marginBottom).not.toBe("");
      expect(style.marginLeft).not.toBe("");
    }
    // EMU margins are converted: 0.25in === 24px at 96dpi.
    expect(Number.parseFloat((paragraphs[0] as HTMLElement).style.marginLeft)).toBeCloseTo(24, 3);
    // spaceAfter is in points and is written to margin-bottom (8pt === 10.67px).
    expect(Number.parseFloat((paragraphs[1] as HTMLElement).style.marginBottom)).toBeCloseTo(8 * (96 / 72), 3);
  });

  it("maps group child coordinates through nested transforms", () => {
    const { container } = renderSlide({
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
            children: sp({ x: inch(2), y: inch(2), cx: inch(1), cy: inch(1), text: paragraph({ runs: run({ text: "inside" }) }) }),
          }),
        },
      ],
    });
    const inner = container.querySelector("[data-shape-kind='shape']") as HTMLElement;
    expect(inner).not.toBeNull();
    // Group scale halves the child space.
    const scaled = inner.parentElement as HTMLElement;
    expect(scaled.style.transform).toContain("scale(0.5, 0.5)");
  });

  it("draws custom geometry paths", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            customGeometryPath:
              '<a:path w="200" h="200"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:cubicBezTo><a:pt x="50" y="0"/><a:pt x="150" y="200"/><a:pt x="200" y="200"/></a:cubicBezTo><a:close/></a:path>',
            fill: '<a:solidFill><a:srgbClr val="00AA00"/></a:solidFill>',
            text: paragraph({ runs: run({ text: "curve" }) }),
          }),
        },
      ],
    });
    const d = container.querySelector("path")?.getAttribute("d") ?? "";
    expect(d).toContain("C");
    expect(d).toContain("Z");
  });

  it("reports unsupported graphic frames as labelled placeholders", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: `<p:graphicFrame>
            <p:nvGraphicFramePr><p:cNvPr id="7" name="Diagram"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
            <p:xfrm><a:off x="0" y="0"/><a:ext cx="${inch(4)}" cy="${inch(2)}"/></p:xfrm>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
              <dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId1"/>
            </a:graphicData></a:graphic>
          </p:graphicFrame>`,
        },
      ],
    });
    const placeholder = container.querySelector("[data-unsupported]");
    expect(placeholder?.textContent).toContain("SmartArt");
  });

  it("replaces an image that fails to decode and reports it", () => {
    const warnings: string[] = [];
    const parsed = parsePresentationBytes(
      buildPptx({
        slides: [
          {
            part: "ppt/slides/slide1.xml",
            shapes: pic({ rId: "rId2", cx: inch(2), cy: inch(2) }),
            rels: { rId2: { target: "../media/image1.png", type: `${NS.r}/image` } },
          },
        ],
        // Sniffs as a PNG (valid signature) but is not decodable in a browser.
        media: [{ part: "ppt/media/image1.png", bytes: RED_PNG.subarray(0, 16) }],
      }),
    );
    const { container } = render(
      <SlideView
        slide={parsed.model.slides[0]!}
        media={parsed.media}
        width={1280}
        height={720}
        scope="decode"
        interactive
        warn={(code) => warnings.push(code)}
      />,
    );
    const image = container.querySelector("img");
    expect(image).not.toBeNull();
    fireEvent.error(image as HTMLImageElement);
    expect(container.querySelector("[data-placeholder='image-decode']")).not.toBeNull();
    expect(warnings).toContain("picture.decode");
  });

  it("labels a genuinely empty slide instead of leaving it blank", () => {
    const { container } = renderSlide({
      slides: [{ part: "ppt/slides/slide1.xml", shapes: "" }],
    });
    const notice = container.querySelector(".radv-empty-slide");
    expect(notice).not.toBeNull();
    expect(notice?.textContent).toBeTruthy();
  });

  it("renders hyperlinks safely and never emits javascript: URLs", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            text: paragraph({
              runs: run({ text: "click", linkRId: "rId3" }),
            }),
          }),
          rels: {
            rId3: { target: "https://example.com/report", type: `${NS.r}/hyperlink`, external: true },
          },
        },
      ],
    });
    const anchor = container.querySelector("a");
    expect(anchor?.getAttribute("href")).toBe("https://example.com/report");
    expect(anchor?.getAttribute("rel")).toBe("noopener noreferrer");
  });

  it("drops javascript: hyperlinks", () => {
    const { container } = renderSlide({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: sp({
            text: paragraph({ runs: run({ text: "evil", linkRId: "rId3" }) }),
          }),
          rels: {
            rId3: { target: "javascript:alert(document.cookie)", type: `${NS.r}/hyperlink`, external: true },
          },
        },
      ],
    });
    expect(container.querySelector("a")).toBeNull();
    expect(container.textContent).toContain("evil");
  });

  it("does not create links when rendering non-interactively", () => {
    const parsed = parsePresentationBytes(
      buildPptx({
        slides: [
          {
            part: "ppt/slides/slide1.xml",
            shapes: sp({ text: paragraph({ runs: run({ text: "link", linkRId: "rId3" }) }) }),
            rels: {
              rId3: { target: "https://example.com", type: `${NS.r}/hyperlink`, external: true },
            },
          },
        ],
      }),
    );
    const { container } = render(
      <SlideView
        slide={parsed.model.slides[0]!}
        media={parsed.media}
        width={1280}
        height={720}
        scope="thumb"
        interactive={false}
      />,
    );
    expect(container.querySelector("a")).toBeNull();
    expect(container.textContent).toContain("link");
  });
});
