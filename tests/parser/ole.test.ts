import { describe, expect, it } from "vitest";
import { OLE_LABELS, hasOleHeader, sniffOleKind } from "../../src/parser/ole";
import { readArchive } from "../../src/parser/archive";
import { detectFormat } from "../../src/parser/document";
import { normalizeError } from "../../src/parser/loadPresentation";

const OLE_HEADER = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** A compound file with the given stream names stored as UTF-16LE, as CFB does. */
function compoundFile(names: string[], padding = 24): Uint8Array {
  const bytes: number[] = [...OLE_HEADER];
  for (let i = 0; i < padding; i += 1) bytes.push(0);
  for (const name of names) {
    for (let i = 0; i < name.length; i += 1) {
      bytes.push(name.charCodeAt(i) & 0xff, 0);
    }
    for (let i = 0; i < 8; i += 1) bytes.push(0);
  }
  return new Uint8Array(bytes);
}

describe("sniffOleKind", () => {
  it("requires the compound-file signature", () => {
    expect(hasOleHeader(compoundFile([]))).toBe(true);
    expect(hasOleHeader(new Uint8Array([0xd0, 0xcf, 0x11]))).toBe(false);
    expect(sniffOleKind(new Uint8Array([1, 2, 3, 4]))).toBe("unknown");
  });

  it("identifies each legacy binary family by its stream names", () => {
    expect(sniffOleKind(compoundFile(["Workbook"]))).toBe("excel");
    expect(sniffOleKind(compoundFile(["Book"]))).toBe("excel");
    expect(sniffOleKind(compoundFile(["WordDocument"]))).toBe("word");
    expect(sniffOleKind(compoundFile(["PowerPoint Document"]))).toBe("powerpoint");
  });

  it("identifies an encrypted package from its streams", () => {
    expect(sniffOleKind(compoundFile(["EncryptionInfo", "EncryptedPackage"]))).toBe("encrypted");
  });

  it("says unknown rather than guessing when no stream name is found", () => {
    expect(sniffOleKind(compoundFile(["CompObj", "SummaryInformation"]))).toBe("unknown");
    expect(OLE_LABELS.unknown).toBeNull();
  });

  it("reports the encrypted package as encrypted, not as legacy", () => {
    const encrypted = compoundFile(["EncryptedPackage"]);
    expect(() => readArchive(encrypted)).toThrowError(expect.objectContaining({ code: "ENCRYPTED" }));
  });

  it("reports each legacy binary format by name instead of calling it encrypted", () => {
    const legacy = compoundFile(["Workbook"]);
    let thrown: unknown = null;
    try {
      readArchive(legacy);
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toMatchObject({ code: "LEGACY_BINARY" });
    expect((thrown as Error).message).toContain(".xls");

    // The public error model treats it as a format the viewer does not render,
    // which is what a user needs to be told, rather than a load failure.
    const normalized = normalizeError(thrown);
    expect(normalized.code).toBe("UNSUPPORTED_FORMAT");
    expect(normalized.message).toContain(".xlsx");
    expect(normalized.message).not.toMatch(/encrypted/i);
  });

  it("still calls an unidentifiable compound file encrypted", () => {
    // The conservative wording is kept when the streams cannot be read, which is
    // also what a bare OLE signature (nothing but the header) looks like.
    const bare = new Uint8Array([...OLE_HEADER, 0, 0, 0, 0]);
    expect(() => readArchive(bare)).toThrowError(expect.objectContaining({ code: "ENCRYPTED" }));
  });

  it("keeps detecting the container as legacy for format detection", () => {
    expect(detectFormat(compoundFile(["Workbook"]))).toBe("legacy");
  });
});
