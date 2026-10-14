/**
 * PDF test fixture builder.
 *
 * Writes a small, valid, uncompressed PDF by hand — a catalog, a page tree,
 * Helvetica text and a correct cross-reference table — so the suite can exercise
 * real PDF parsing without checking in a binary or pulling in a PDF writer.
 */

export interface PdfPageSpec {
  text?: string;
  /** MediaBox size in PDF points; defaults to US Letter */
  width?: number;
  height?: number;
}

export interface PdfFixtureOptions {
  pages?: (string | PdfPageSpec)[];
  title?: string;
  author?: string;
  /** omit the trailer so the file is recognisably a broken PDF */
  broken?: boolean;
}

function escapePdfText(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

export function buildPdf(options: PdfFixtureOptions = {}): Uint8Array {
  const specs: PdfPageSpec[] = (options.pages ?? ["Hello page one", "Hello page two"]).map((page) =>
    typeof page === "string" ? { text: page } : page,
  );
  const objects: string[] = [];
  const add = (body: string): number => {
    objects.push(body);
    return objects.length;
  };

  const fontNumber = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  const pagesNumber = add(""); // filled in once the page numbers are known
  const kids: number[] = [];

  for (const spec of specs) {
    const content = `BT /F1 24 Tf 72 700 Td (${escapePdfText(spec.text ?? "")}) Tj ET`;
    const contentNumber = add(`<< /Length ${content.length} >>\nstream\n${content}\nendstream`);
    const width = spec.width ?? 612;
    const height = spec.height ?? 792;
    kids.push(
      add(
        `<< /Type /Page /Parent ${pagesNumber} 0 R /MediaBox [0 0 ${width} ${height}] ` +
          `/Resources << /Font << /F1 ${fontNumber} 0 R >> >> /Contents ${contentNumber} 0 R >>`,
      ),
    );
  }

  objects[pagesNumber - 1] =
    `<< /Type /Pages /Kids [${kids.map((number) => `${number} 0 R`).join(" ")}] /Count ${kids.length} >>`;
  const catalogNumber = add(`<< /Type /Catalog /Pages ${pagesNumber} 0 R >>`);

  const infoParts: string[] = [];
  if (options.title) infoParts.push(`/Title (${escapePdfText(options.title)})`);
  if (options.author) infoParts.push(`/Author (${escapePdfText(options.author)})`);
  const infoNumber = infoParts.length > 0 ? add(`<< ${infoParts.join(" ")} >>`) : null;

  let out = "%PDF-1.4\n";
  const offsets: number[] = [];
  for (let i = 0; i < objects.length; i += 1) {
    offsets.push(out.length);
    out += `${i + 1} 0 obj\n${objects[i]}\nendobj\n`;
  }

  if (options.broken) return new TextEncoder().encode(`${out}%%EOF\n`);

  const xrefStart = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets) out += `${String(offset).padStart(10, "0")} 00000 n \n`;
  out +=
    `trailer\n<< /Size ${objects.length + 1} /Root ${catalogNumber} 0 R` +
    `${infoNumber ? ` /Info ${infoNumber} 0 R` : ""} >>\nstartxref\n${xrefStart}\n%%EOF\n`;
  return new TextEncoder().encode(out);
}

/** A PDF with the given page texts, two pages by default. */
export function buildTextPdf(...pages: string[]): Uint8Array {
  return buildPdf({ pages });
}

/** Not a PDF at all, but it starts with the PDF signature. */
export function buildFakePdf(): Uint8Array {
  return new TextEncoder().encode("%PDF-1.4\nthis is not a pdf body at all\n");
}
