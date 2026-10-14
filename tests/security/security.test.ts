import { describe, expect, it } from "vitest";
import { zipSync } from "fflate";
import {
  ArchiveError,
  DEFAULT_ARCHIVE_LIMITS,
  isSafeEntryName,
  isUsefulPart,
  readArchive,
  sniffImageKind,
} from "../../src/parser/archive";
import { XmlError, parseXml } from "../../src/parser/xml";
import { sanitizeUrl } from "../../src/utils/sanitize";
import { parsePresentationBytes } from "../../src/parser/presentation";
import { buildPptx, NS, paragraph, run, sp, TINY_PNG } from "../fixtures/pptx";

const encoder = new TextEncoder();

function bytes(text: string): Uint8Array {
  return encoder.encode(text);
}

describe("archive limits", () => {
  it("rejects input that is not a ZIP package", () => {
    expect(() => readArchive(bytes("this is not a zip file at all"))).toThrowError(
      expect.objectContaining({ code: "NOT_A_ZIP" }),
    );
  });

  it("rejects empty input", () => {
    expect(() => readArchive(new Uint8Array(0))).toThrowError(
      expect.objectContaining({ code: "EMPTY_ARCHIVE" }),
    );
  });

  it("detects encrypted / legacy binary Office files", () => {
    const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]);
    expect(() => readArchive(ole)).toThrowError(expect.objectContaining({ code: "ENCRYPTED" }));
  });

  it("enforces the entry-count limit", () => {
    const buffer = buildPptx({
      slides: [{ part: "ppt/slides/slide1.xml", shapes: sp({ text: paragraph({ runs: run({ text: "x" }) }) }) }],
    });
    expect(() => readArchive(buffer, { maxEntries: 3 })).toThrowError(
      expect.objectContaining({ code: "ARCHIVE_LIMIT_EXCEEDED" }),
    );
  });

  it("enforces the per-entry size limit", () => {
    const buffer = buildPptx({
      slides: [{ part: "ppt/slides/slide1.xml", shapes: sp({ text: paragraph({ runs: run({ text: "x" }) }) }) }],
    });
    expect(() => readArchive(buffer, { maxEntryBytes: 512 })).toThrowError(
      expect.objectContaining({ code: "ARCHIVE_LIMIT_EXCEEDED" }),
    );
  });

  it("enforces the total uncompressed size limit", () => {
    const buffer = buildPptx({
      slides: [{ part: "ppt/slides/slide1.xml", shapes: sp({ text: paragraph({ runs: run({ text: "x" }) }) }) }],
    });
    expect(() => readArchive(buffer, { maxTotalBytes: 4096 })).toThrowError(
      expect.objectContaining({ code: "ARCHIVE_LIMIT_EXCEEDED" }),
    );
  });

  it("enforces the input size limit", () => {
    const buffer = buildPptx({
      slides: [{ part: "ppt/slides/slide1.xml", shapes: sp({ text: paragraph({ runs: run({ text: "x" }) }) }) }],
    });
    expect(() => readArchive(buffer, { maxCompressedBytes: 64 })).toThrowError(
      expect.objectContaining({ code: "ARCHIVE_LIMIT_EXCEEDED" }),
    );
  });

  it("rejects a decompression bomb by expansion ratio", () => {
    // 6 MiB of zeros compresses to a few KiB, which is far above the default
    // 200:1 ratio threshold.
    const bomb = zipSync({ "ppt/media/big.bin": new Uint8Array(6 * 1024 * 1024) }, { level: 9 });
    expect(bomb.byteLength / (6 * 1024 * 1024)).toBeLessThan(1 / DEFAULT_ARCHIVE_LIMITS.maxCompressionRatio);
    expect(() => readArchive(bomb)).toThrowError(
      expect.objectContaining({ code: "ARCHIVE_LIMIT_EXCEEDED" }),
    );
  });

  it("never inflates parts the renderer does not need", () => {
    const buffer = zipSync({
      "[Content_Types].xml": bytes("<Types/>"),
      "ppt/embeddings/huge.xlsx": new Uint8Array(2 * 1024 * 1024),
      "docProps/thumbnail.jpeg": new Uint8Array(1024),
      "ppt/media/kept.png": TINY_PNG,
    });
    // The ratio guard is unrelated to this test, so it is relaxed here.
    const archive = readArchive(buffer, { maxCompressionRatio: 100_000 });
    expect(archive.entries.has("ppt/embeddings/huge.xlsx")).toBe(false);
    expect(archive.entries.has("docprops/thumbnail.jpeg")).toBe(false);
    expect(archive.entries.has("ppt/media/kept.png")).toBe(true);
  });

  it("flags archive errors with a machine-readable code", () => {
    try {
      readArchive(bytes("nope"));
      throw new Error("expected a throw");
    } catch (error) {
      expect(error).toBeInstanceOf(ArchiveError);
      expect((error as ArchiveError).code).toBe("NOT_A_ZIP");
    }
  });
});

describe("entry name safety", () => {
  it("rejects traversal, absolute and control-character names", () => {
    expect(isSafeEntryName("ppt/slides/slide1.xml")).toBe(true);
    expect(isSafeEntryName("../../etc/passwd")).toBe(false);
    expect(isSafeEntryName("/etc/passwd")).toBe(false);
    expect(isSafeEntryName("ppt\\slides\\slide1.xml")).toBe(false);
    expect(isSafeEntryName("ppt/slides/\u0000evil.xml")).toBe(false);
    expect(isSafeEntryName("a".repeat(600))).toBe(false);
    expect(isSafeEntryName("")).toBe(false);
  });

  it("ignores unsafe entries even when they are inside a real package", () => {
    const buffer = zipSync({
      "[Content_Types].xml": bytes("<Types/>"),
      "../../escape.xml": bytes("<x/>"),
      "ppt/slides/../../escape2.xml": bytes("<x/>"),
      "ppt/slides/slide1.xml": bytes("<p:sld/>"),
    });
    const archive = readArchive(buffer);
    expect(archive.names).toContain("../../escape.xml");
    expect(archive.entries.has("../../escape.xml")).toBe(false);
    expect(archive.entries.has("ppt/slides/slide1.xml")).toBe(true);
  });

  it("only treats presentation parts as useful", () => {
    expect(isUsefulPart("[Content_Types].xml")).toBe(true);
    expect(isUsefulPart("ppt/slides/_rels/slide1.xml.rels")).toBe(true);
    expect(isUsefulPart("ppt/media/image1.png")).toBe(true);
    expect(isUsefulPart("ppt/embeddings/book1.xlsx")).toBe(false);
    expect(isUsefulPart("ppt/notesSlides/notesSlide1.xml")).toBe(false);
  });
});

describe("hostile XML", () => {
  it("refuses documents with a DOCTYPE or ENTITY declaration", () => {
    const billionLaughs = `<?xml version="1.0"?>
<!DOCTYPE lolz [
  <!ENTITY lol "lol">
  <!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
]>
<lolz>&lol2;</lolz>`;
    expect(() => parseXml(billionLaughs)).toThrowError(XmlError);
  });

  it("does not expand entities and decodes only the standard five", () => {
    const root = parseXml(`<a:t>R&amp;D &lt;tag&gt; &#65; &#x42;</a:t>`);
    expect(root?.text).toBe("R&D <tag> A B");
  });

  it("keeps an unknown entity literal instead of expanding it", () => {
    const root = parseXml(`<a:t>&notAnEntity;</a:t>`);
    expect(root?.text).toBe("&notAnEntity;");
  });

  it("reports unparseable XML with the offending part name", () => {
    expect(() => parseXml("<a:t>value</a:t", { part: "ppt/slides/slide1.xml" })).toThrowError(
      /ppt\/slides\/slide1\.xml/,
    );
  });

  it("documents the parser's leniency for unclosed tags", () => {
    // fast-xml-parser recovers from unclosed and mismatched tags instead of
    // throwing. That is deliberate: a truncated part degrades to a readable
    // (if incomplete) tree plus diagnostics rather than failing the whole deck.
    const recovered = parseXml("<p:sp><a:t>value</a:t>");
    expect(recovered?.name).toBe("p:sp");
    // A mismatched close tag is recovered as text on the opening element.
    expect(parseXml("<a:t>value</a:b>")?.text).toBe("value");
  });

  it("survives a hostile slide part by skipping it with a diagnostic", () => {
    const buffer = buildPptx({
      slides: [{ part: "ppt/slides/slide1.xml", shapes: "" }],
      extraParts: {
        "ppt/slides/slide1.xml": `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e "boom">]><p:sld/>`,
      },
    });
    const parsed = parsePresentationBytes(buffer);
    expect(parsed.model.slides).toHaveLength(1);
    expect(parsed.model.diagnostics.some((d) => d.code === "part.unreadable")).toBe(true);
  });

  it("ignores an XML declaration that would otherwise become the root", () => {
    const xml = `<?xml version="1.0"?><p:presentation xmlns:p="u"><child/></p:presentation>`;
    expect(parseXml(xml)?.name).toBe("p:presentation");
  });
});

describe("link sanitisation", () => {
  it("allows http, https, mailto and tel", () => {
    expect(sanitizeUrl("https://example.com/a?b=1").safe).toBe(true);
    expect(sanitizeUrl("http://example.com").safe).toBe(true);
    expect(sanitizeUrl("mailto:team@example.com").safe).toBe(true);
    expect(sanitizeUrl("tel:+15551234").safe).toBe(true);
  });

  it("blocks script, data and file schemes", () => {
    for (const url of [
      "javascript:alert(1)",
      "JavaScript:alert(1)",
      "  javascript:alert(1)",
      "data:text/html;base64,PHNjcmlwdD4=",
      "vbscript:msgbox(1)",
      "file:///etc/passwd",
      "about:blank",
    ]) {
      const result = sanitizeUrl(url);
      expect(result.safe, url).toBe(false);
      expect(result.href, url).toBe("");
    }
  });

  it("blocks credentials embedded in an otherwise valid URL", () => {
    expect(sanitizeUrl("https://user:pass@example.com/secret").safe).toBe(false);
  });

  it("rejects relative and over-long targets by default", () => {
    expect(sanitizeUrl("/relative/path").safe).toBe(false);
    expect(sanitizeUrl(`https://example.com/${"a".repeat(3000)}`).safe).toBe(false);
  });
});

describe("media sniffing", () => {
  it("identifies formats from magic bytes rather than extensions", () => {
    expect(sniffImageKind(TINY_PNG)).toBe("png");
    expect(sniffImageKind(new Uint8Array([0xff, 0xd8, 0xff, 0xe0]))).toBe("jpeg");
    expect(sniffImageKind(bytes("GIF89a"))).toBe("gif");
    expect(sniffImageKind(bytes("<svg xmlns='http://www.w3.org/2000/svg'></svg>"))).toBe("svg");
    expect(sniffImageKind(new Uint8Array([0xd7, 0xcd, 0xc6, 0x9a]))).toBe("wmf");
    expect(sniffImageKind(bytes("not an image"))).toBe("unknown");
  });

  it("exposes a data URL when object URLs are unavailable", () => {
    const buffer = buildPptx({
      slides: [
        {
          part: "ppt/slides/slide1.xml",
          shapes: `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`,
          rels: { rId2: { target: "../media/image1.png", type: `${NS.r}/image` } },
        },
      ],
      media: [{ part: "ppt/media/image1.png", bytes: TINY_PNG }],
    });
    // The store decides its strategy when it is created, so the capability is
    // removed before parsing, exactly as it would be when rendering on a server.
    const originalCreate = URL.createObjectURL;
    Object.defineProperty(URL, "createObjectURL", { configurable: true, writable: true, value: undefined });
    try {
      const parsed = parsePresentationBytes(buffer);
      expect(parsed.media.url("ppt/media/image1.png")).toMatch(/^data:image\/png;base64,/);
    } finally {
      Object.defineProperty(URL, "createObjectURL", {
        configurable: true,
        writable: true,
        value: originalCreate,
      });
    }
  });
});
