import { describe, expect, it } from "vitest";
import { renderToStaticMarkup } from "react-dom/server";
import { readImageInfo, parseImageDocument } from "../../src/parser/image";
import { readTiffPages } from "../../src/parser/tiff";
import { detectFormat } from "../../src/parser/document";
import { mergeLabels } from "../../src/components/labels";
import { brokenTiff, buildTiff } from "../fixtures/tiff";

describe("TIFF", () => {
  it("reads every page and its natural size from the IFD chain", () => {
    const tiff = buildTiff({
      pages: [
        { width: 320, height: 200 },
        { width: 64, height: 48, rgb: [0, 128, 255] },
      ],
    });
    expect(readTiffPages(tiff)).toEqual([
      { width: 320, height: 200 },
      { width: 64, height: 48 },
    ]);
    expect(readImageInfo(tiff)).toEqual({ kind: "tiff", width: 320, height: 200 });
  });

  it("is detected from its bytes and becomes a multi-page image document", () => {
    const tiff = buildTiff({
      pages: [
        { width: 320, height: 200 },
        { width: 320, height: 200 },
      ],
    });
    expect(detectFormat(tiff)).toBe("image");

    const parsed = parseImageDocument(tiff);
    expect(parsed.format).toBe("image");
    expect(parsed.model.slides).toHaveLength(2);
    expect(parsed.model.width).toBe(Math.round((320 / 96) * 914_400));
    expect(parsed.model.height).toBe(Math.round((200 / 96) * 914_400));
    expect(parsed.model.slides[1]?.name).toBe("Page 2");
    expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain("tiff.multi-page");
    parsed.dispose();
  });

  it("renders a TIFF page through the canvas renderer, not an img", () => {
    const parsed = parseImageDocument(buildTiff({ pages: [{ width: 20, height: 10 }] }));
    const page = parsed.model.slides[0]!;
    const element = parsed.renderPage?.({
      page,
      width: 20,
      height: 10,
      scale: 1,
      interactive: true,
      media: parsed.media,
      labels: mergeLabels(undefined),
      warn: () => undefined,
    });
    expect(element).toBeTruthy();
    const markup = renderToStaticMarkup(element as React.ReactElement);
    expect(markup).toContain("rv-tiff-page");
    expect(markup).toContain("radv-tiff-page");
    expect(markup).not.toContain("<img");
    parsed.dispose();
  });

  it("falls back to a default page when the directory is unreadable", () => {
    const parsed = parseImageDocument(brokenTiff());
    expect(parsed.model.slides).toHaveLength(1);
    expect(parsed.model.width).toBe(Math.round((1024 / 96) * 914_400));
    expect(parsed.model.diagnostics.map((diagnostic) => diagnostic.code)).toContain("tiff.directory");
    parsed.dispose();
  });

  it("optimistically decodes the generated fixture with the TIFF decoder", async () => {
    // This is the same decoder path the browser renderer takes. It proves the
    // fixture is a real TIFF, not just bytes the header parser happens to accept.
    const UTIF = await import("utif");
    const pages = buildTiff({
      pages: [
        { width: 4, height: 3 },
        { width: 2, height: 2, rgb: [0, 128, 255] },
      ],
    });
    const buffer = pages.buffer.slice(pages.byteOffset, pages.byteOffset + pages.byteLength) as ArrayBuffer;
    const directories = UTIF.decode(buffer);
    expect(directories).toHaveLength(2);
    const second = directories[1] as {
      t256: number[];
      t257: number[];
      width?: number;
      height?: number;
    };
    // UTIF keeps raw tags until `decodeImage` materialises width/height.
    expect(second.t256[0]).toBe(2);
    expect(second.t257[0]).toBe(2);
    UTIF.decodeImage(buffer, second);
    expect(second.width).toBe(2);
    expect(second.height).toBe(2);
    const rgba = UTIF.toRGBA8(second);
    expect(rgba).toHaveLength(2 * 2 * 4);
    expect([...rgba.slice(0, 4)]).toEqual([0, 128, 255, 255]);
  });
});
