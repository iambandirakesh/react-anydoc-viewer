/**
 * OLE/CFB (compound file) identification.
 *
 * A `.doc`, `.ppt` and `.xls` from Office 97-2003, and an encrypted OOXML file,
 * all start with the same eight signature bytes. They are routinely reported as
 * one thing ("encrypted"), which is a misleading thing to tell a user who has
 * simply opened an old spreadsheet.
 *
 * A compound file keeps its streams in a directory whose names are stored as
 * UTF-16LE, so the two families can be told apart from the bytes:
 *
 *   EncryptedPackage / EncryptionInfo   an encrypted (password-protected) package
 *   Workbook / Book                     a legacy Excel workbook (.xls, BIFF5/8)
 *   WordDocument                        a legacy Word document (.doc)
 *   PowerPoint Document                 a legacy PowerPoint presentation (.ppt)
 *
 * A bounded scan is used because the directory can sit anywhere in the file once
 * its sectors have been extended, and because the input is untrusted: nothing
 * beyond {@link MAX_SCAN_BYTES} is ever examined, and the caller keeps its own
 * size limits. Not finding a name is not evidence of anything, so the answer is
 * then "unknown" and the caller falls back to its more conservative wording.
 */

export type OleKind =
  /** an encrypted or password-protected OOXML package */
  | "encrypted"
  /** a legacy binary Excel workbook */
  | "excel"
  /** a legacy binary Word document */
  | "word"
  /** a legacy binary PowerPoint presentation */
  | "powerpoint"
  /** a compound file whose streams could not be identified */
  | "unknown";

/** How much of the file is searched for directory names. */
const MAX_SCAN_BYTES = 8 * 1024 * 1024;

const OLE_HEADER = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export function hasOleHeader(bytes: Uint8Array): boolean {
  if (bytes.length < OLE_HEADER.length) return false;
  for (let i = 0; i < OLE_HEADER.length; i += 1) {
    if (bytes[i] !== OLE_HEADER[i]) return false;
  }
  return true;
}

/** UTF-16LE code units of an ASCII stream name. */
function utf16leBytes(name: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < name.length; i += 1) {
    out.push(name.charCodeAt(i) & 0xff, 0);
  }
  return out;
}

function findSequence(bytes: Uint8Array, needle: readonly number[], limit: number): boolean {
  const end = Math.min(bytes.length, limit);
  const first = needle[0] as number;
  for (let i = 0; i + needle.length <= end; i += 1) {
    if (bytes[i] !== first) continue;
    let matched = true;
    for (let j = 1; j < needle.length; j += 1) {
      if (bytes[i + j] !== needle[j]) {
        matched = false;
        break;
      }
    }
    if (matched) return true;
  }
  return false;
}

function asciiBytes(name: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < name.length; i += 1) out.push(name.charCodeAt(i) & 0xff);
  return out;
}

/**
 * Which family a compound file belongs to.
 *
 * UTF-16LE (how CFB stores names) is looked for first; a plain-ASCII fallback
 * covers files from tools that write the directory differently.
 */
export function sniffOleKind(bytes: Uint8Array): OleKind {
  if (!hasOleHeader(bytes)) return "unknown";
  const limit = Math.min(bytes.length, MAX_SCAN_BYTES);

  const has = (name: string): boolean => {
    const utf16 = utf16leBytes(name);
    if (findSequence(bytes, utf16, limit)) return true;
    return findSequence(bytes, asciiBytes(name), limit);
  };

  if (has("EncryptedPackage") || has("EncryptionInfo")) return "encrypted";
  if (has("WordDocument")) return "word";
  if (has("PowerPoint Document")) return "powerpoint";
  // BIFF8 names the stream "Workbook"; BIFF5 used "Book".
  if (has("Workbook") || has("Book")) return "excel";
  return "unknown";
}

/** How a compound file is described in an error message. */
export const OLE_LABELS: Record<OleKind, string | null> = {
  encrypted: "an encrypted or password-protected Office file",
  excel: "a legacy Excel 97-2003 workbook (.xls)",
  word: "a legacy Word 97-2003 document (.doc)",
  powerpoint: "a legacy PowerPoint 97-2003 presentation (.ppt)",
  unknown: null,
};
