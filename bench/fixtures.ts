/**
 * Realistic, large fixtures for the load benchmark.
 *
 * The suite's own fixtures are deliberately small (a handful of slides), so they
 * say nothing about how the viewer behaves on the files a user actually opens.
 * These generators scale the same builders up — a forty-slide deck with media, a
 * long Word document with tables, a hundred-page PDF — and re-compress the result
 * at the level Office uses, because an uncompressed package would make the
 * archive reader look faster than it is.
 */

import { unzipSync, zipSync } from "fflate";
import { buildPptx, bulletChar, inch, paragraph, pic, run, sp, type SlideSpec } from "../tests/fixtures/pptx";
import { buildDocx, wCell, wParagraph, wRun, wTable } from "../tests/fixtures/docx";
import { buildPdf } from "../tests/fixtures/pdf";

/**
 * 24-bit BMP with noise in the pixels.
 *
 * The noise matters: a smooth gradient deflates to a few kilobytes, which would
 * make a "5 MB deck" fixture a 60 KB file the moment it is zipped. A linear
 * congruential sequence is unpredictable enough to stay large.
 */
function noisyBmp(width: number, height: number, seed = 1): Uint8Array {
  const rowSize = Math.ceil((width * 3) / 4) * 4;
  const pixels = rowSize * height;
  const out = new Uint8Array(54 + pixels);
  const view = new DataView(out.buffer);
  out[0] = 0x42;
  out[1] = 0x4d;
  view.setUint32(2, 54 + pixels, true);
  view.setUint32(10, 54, true);
  view.setUint32(14, 40, true);
  view.setInt32(18, width, true);
  view.setInt32(22, height, true);
  view.setUint16(26, 1, true);
  view.setUint16(28, 24, true);
  view.setUint32(34, pixels, true);
  view.setInt32(38, 2835, true);
  view.setInt32(42, 2835, true);
  // xorshift32: exact in 32-bit integer arithmetic (a multiply-based generator
  // loses its low bits to double precision, and the "noise" deflates to nothing).
  let state = (seed | 1) >>> 0;
  for (let i = 54; i < out.length; i += 1) {
    state ^= state << 13;
    state >>>= 0;
    state ^= state >>> 17;
    state ^= state << 5;
    state >>>= 0;
    out[i] = state & 0xff;
  }
  return out;
}

/** Re-zip at Office's compression level so the archive path is exercised fairly. */
function recompress(bytes: Uint8Array): Uint8Array {
  return zipSync(unzipSync(bytes), { level: 6 });
}

const IMAGE_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const LOREM =
  "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor " +
  "incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud.";

const WORDS = [
  "alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel",
  "india", "juliet", "kilo", "lima", "mike", "november", "oscar", "papa",
  "quebec", "romeo", "sierra", "tango", "uniform", "victor", "whiskey", "xray",
];

/**
 * A deterministic pseudo-sentence, distinct per seed.
 *
 * Padded text is the enemy of an honest Word fixture: twelve hundred copies of one
 * paragraph deflate to a few kilobytes, so the archive reader would be measured on
 * a document nobody would ever open.
 */
function sentence(seed: number, words = 24): string {
  const out: string[] = [];
  for (let index = 0; index < words; index += 1) {
    out.push(WORDS[(seed * 31 + index * 7 + ((seed * index) % 5)) % WORDS.length]!);
  }
  // A short unique token per paragraph, so twelve hundred paragraphs do not
  // deflate to a few kilobytes the way twelve hundred identical ones would.
  const token = ((seed * 2654435761) >>> 0).toString(36);
  return `${out.join(" ")} [${token}].`;
}

/**
 * A deck of `slides` slides: title, bullets, notes, a rectangle, and pictures.
 *
 * With `heavyMedia` every slide carries two pictures of its own, which is what an
 * image-led deck looks like: the parser still only reads the parts it needs, but
 * the browser then has to decode whatever is on screen.
 */
export function bigDeck(slides = 40, heavyMedia = false): Uint8Array {
  const images = heavyMedia
    ? Array.from({ length: slides * 2 }, (_, index) => ({
        part: `ppt/media/image${index + 1}.bmp`,
        bytes: noisyBmp(200, 150, index * 2 + 1),
      }))
    : [
        { part: "ppt/media/image1.bmp", bytes: noisyBmp(700, 500, 11) },
        { part: "ppt/media/image2.bmp", bytes: noisyBmp(700, 500, 29) },
        { part: "ppt/media/image3.bmp", bytes: noisyBmp(700, 500, 47) },
      ];

  const specs: SlideSpec[] = [];
  for (let index = 0; index < slides; index += 1) {
    const shapes: string[] = [
      sp({
        id: 2,
        name: "Title",
        ph: "title",
        x: inch(0.6),
        y: inch(0.4),
        cx: inch(12),
        cy: inch(1.1),
        text: paragraph({
          runs: run({ text: `Slide ${index + 1}: quarterly review`, size: 3200, bold: true }),
        }),
      }),
    ];

    const bullets = Array.from({ length: 5 }, (_, line) =>
      paragraph({
        level: 0,
        bullet: bulletChar(),
        runs: run({ text: `${line + 1}. ${LOREM.slice(0, 120)}`, size: 1800 }),
      }),
    ).join("");
    shapes.push(
      sp({
        id: 3,
        name: "Body",
        ph: "body",
        x: inch(0.6),
        y: inch(1.7),
        cx: inch(6.4),
        cy: inch(4.6),
        text: bullets,
      }),
    );

    // A second text box and a plain rectangle, so the shape tree is not uniform.
    shapes.push(
      sp({
        id: 4,
        name: "Notes",
        txBox: true,
        x: inch(7.4),
        y: inch(1.7),
        cx: inch(5.2),
        cy: inch(2),
        text: paragraph({ runs: run({ text: LOREM, size: 1400 }) }),
      }),
      sp({ id: 5, name: "Bar", preset: "roundRect", x: inch(7.4), y: inch(4), cx: inch(5.2), cy: inch(2.3) }),
    );

    const slideImages = heavyMedia
      ? [images[index * 2]!, images[index * 2 + 1]!]
      : index % 5 === 0
        ? [images[index % images.length]!]
        : [];
    slideImages.forEach((image, position) => {
      shapes.push(
        pic({
          id: 6 + position,
          name: `Picture ${position + 1}`,
          rId: `rId${2 + position}`,
          x: inch(0.6 + position * 3.2),
          y: inch(1.7),
          cx: inch(3),
          cy: inch(2),
        }),
      );
    });

    specs.push({
      part: `ppt/slides/slide${index + 1}.xml`,
      name: `Slide ${index + 1}`,
      shapes: shapes.join(""),
      ...(slideImages.length > 0
        ? {
            rels: Object.fromEntries(
              slideImages.map((image, position) => [
                `rId${2 + position}`,
                { target: `../media/${image.part.replace(/^ppt\/media\//, "")}`, type: IMAGE_REL },
              ]),
            ),
          }
        : {}),
    });
  }

  return recompress(
    buildPptx({
      slides: specs,
      media: images,
      coreTitle: "Benchmark deck",
    }),
  );
}

/** A long document: `paragraphs` body paragraphs with a table every 100. */
export function bigDocx(paragraphs = 1200): Uint8Array {
  const body: string[] = [];
  for (let index = 0; index < paragraphs; index += 1) {
    if (index > 0 && index % 100 === 0) {
      body.push(
        wTable({
          columns: [2000, 2000, 2000],
          borders: true,
          rows: Array.from({ length: 5 }, (_, row) => ({
            cells: Array.from({ length: 3 }, (_unused, column) => ({
              xml: wParagraph({ runs: wRun(sentence(index + row * 10 + column, 8)) }),
            })),
          })),
        }),
      );
    }
    body.push(wParagraph({ runs: wRun(`Paragraph ${index + 1}. ${sentence(index)}`) }));
  }

  return recompress(buildDocx({ body: body.join(""), core: { title: "Benchmark document" } }));
}

/** A PDF of `pages` pages, each with a few lines of body text. */
export function bigPdf(pages = 120): Uint8Array {
  return buildPdf({
    title: "Benchmark PDF",
    pages: Array.from({ length: pages }, (_, page) =>
      Array.from({ length: 12 }, (_, line) => `Page ${page + 1} line ${line + 1}: ${LOREM}`).join("\n"),
    ),
  });
}
