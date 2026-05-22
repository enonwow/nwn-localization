import {
  type LocaleColumn,
  type ParsedTlk,
  type ParsedTlkBundle,
  type TlkBundleConfig,
  type TlkGridRow,
  localeCodeToFieldToken,
  normalizeLocaleCode
} from "./types";

export const TLK_HEADER_SIZE = 20;
export const TLK_ENTRY_SIZE = 40;
export const TLK_TEXT_PRESENT_FLAG = 0x0001;
export const TLK_LOCALE_TO_LANGUAGE_ID: Readonly<Record<string, number>> = {
  EN: 0,
  FR: 1,
  DE: 2,
  IT: 3,
  ES: 4,
  "PT-BR": 4,
  PL: 5
};

const TLK_LANGUAGE_TO_ENCODING: Record<number, string> = {
  0: "windows-1252",
  1: "windows-1252",
  2: "windows-1252",
  3: "windows-1252",
  4: "windows-1252",
  5: "windows-1250"
};

function asArrayBuffer(bufferLike: ArrayBuffer | ArrayBufferView): ArrayBuffer {
  if (bufferLike instanceof ArrayBuffer) return bufferLike;
  const view = new Uint8Array(bufferLike.buffer, bufferLike.byteOffset, bufferLike.byteLength);
  const copy = new Uint8Array(view.byteLength);
  copy.set(view);
  return copy.buffer;
}

export function readAscii(view: DataView, start: number, length: number): string {
  const chars: string[] = [];
  for (let i = 0; i < length; i += 1) {
    chars.push(String.fromCharCode(view.getUint8(start + i)));
  }
  return chars.join("");
}

function decodeBytesSmart(bytes: Uint8Array, languageId: number): string {
  const encodings: string[] = [];
  const preferred = TLK_LANGUAGE_TO_ENCODING[languageId];
  if (preferred) {
    encodings.push(preferred);
  }
  encodings.push("utf-8", "windows-1252");

  let bestText = "";
  let bestPenalty = Number.POSITIVE_INFINITY;

  for (let i = 0; i < encodings.length; i += 1) {
    const encoding = encodings[i];
    try {
      const decoder = new TextDecoder(encoding, { fatal: false });
      const text = decoder.decode(bytes);
      const replacementCount = (text.match(/\uFFFD/g) || []).length;
      const penalty = replacementCount * 1000 + Math.abs(text.length - bytes.length);
      if (penalty < bestPenalty) {
        bestPenalty = penalty;
        bestText = text;
      }
    } catch (error) {
      // Some runtimes (notably older embedded browsers) may not expose legacy encodings.
      // Keep trying fallback decoders when that occurs.
    }
  }

  return bestText;
}

export function fileNameToken(value: string): string {
  return String(value || "").trim().toLowerCase();
}

export function safeFileNameFromPath(value: string): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const normalized = raw.replace(/\\/g, "/");
  const parts = normalized.split("/");
  return parts[parts.length - 1] || "";
}

export function sanitizeSheetCell(value: unknown): string {
  if (value == null) return "";
  return String(value).replace(/\r\n/g, "\n");
}

function normalizeParsedBundle(bundle: TlkBundleConfig): TlkBundleConfig {
  return {
    ...bundle,
    locale: normalizeLocaleCode(bundle.locale),
    dialog: String(bundle.dialog || "").trim(),
    dialogf: String(bundle.dialogf || "").trim(),
    dialogfAuto: Boolean(bundle.dialogfAuto)
  };
}

export function parseSingleTlkBuffer(arrayBufferLike: ArrayBuffer | ArrayBufferView, fileName: string): ParsedTlk {
  const buffer = asArrayBuffer(arrayBufferLike);
  const view = new DataView(buffer);
  if (view.byteLength < TLK_HEADER_SIZE) {
    throw new Error(`${fileName}: file is too small to be a TLK.`);
  }

  const signature = readAscii(view, 0, 4);
  const version = readAscii(view, 4, 4);
  if (signature !== "TLK " || version !== "V3.0") {
    throw new Error(`${fileName}: invalid TLK signature/version (${signature} ${version}).`);
  }

  const languageId = view.getUint32(8, true);
  const entryCount = view.getUint32(12, true);
  const stringBlobOffset = view.getUint32(16, true);
  const expectedDescriptorBytes = entryCount * TLK_ENTRY_SIZE;
  const descriptorsEnd = TLK_HEADER_SIZE + expectedDescriptorBytes;
  if (descriptorsEnd > view.byteLength || stringBlobOffset > view.byteLength) {
    throw new Error(`${fileName}: malformed TLK layout.`);
  }

  const rows = new Array(entryCount);
  for (let i = 0; i < entryCount; i += 1) {
    const base = TLK_HEADER_SIZE + i * TLK_ENTRY_SIZE;
    const flags = view.getUint32(base, true);
    const stringOffset = view.getUint32(base + 28, true);
    const stringLength = view.getUint32(base + 32, true);
    const hasText = (flags & TLK_TEXT_PRESENT_FLAG) !== 0 && stringLength > 0;

    let text = "";
    if (hasText) {
      const textStart = stringBlobOffset + stringOffset;
      const textEnd = textStart + stringLength;
      if (textStart >= 0 && textEnd <= view.byteLength && textStart <= textEnd) {
        const bytes = new Uint8Array(buffer, textStart, stringLength);
        text = decodeBytesSmart(bytes, languageId);
      }
    }
    rows[i] = text;
  }

  return {
    fileName,
    languageId,
    entryCount,
    rows
  };
}

export function buildRowsFromParsedTlkBundles(parsedBundles: readonly ParsedTlkBundle[], localeColumns: readonly LocaleColumn[]): TlkGridRow[] {
  if (!Array.isArray(parsedBundles) || parsedBundles.length === 0) {
    return [];
  }

  const entryCount = parsedBundles[0].parsed.entryCount;
  const rows: TlkGridRow[] = new Array(entryCount);
  const localeColumnsNormalized = localeColumns.length > 0 ? localeColumns : [];
  const englishColumn = localeColumnsNormalized.find(col => col.locale === "EN" && col.variant === "dialog");

  for (let i = 0; i < entryCount; i += 1) {
    const row: TlkGridRow = {
      id: i,
      strRef: i,
      sourceEn: "",
      context: "",
      status: "Draft"
    };

    localeColumnsNormalized.forEach(col => {
      row[col.field] = "";
    });

    parsedBundles.forEach(bundleData => {
      const bundle = normalizeParsedBundle(bundleData.bundle);
      const parsed = bundleData.parsed;
      const localeToken = localeCodeToFieldToken(bundle.locale);
      const dialogField = `loc_${localeToken}`;
      const dialogfField = `loc_${localeToken}_f`;
      row[dialogField] = sanitizeSheetCell(parsed.rows[i] || "");
      if (bundleData.dialogfParsed && !bundle.dialogfAuto) {
        row[dialogfField] = sanitizeSheetCell(bundleData.dialogfParsed.rows[i] || "");
      }
    });

    const sourceFromEN = englishColumn ? row[englishColumn.field] : "";
    if (sourceFromEN) {
      row.sourceEn = String(sourceFromEN);
    } else {
      const firstTextField = localeColumnsNormalized.find(col => String(row[col.field] || "").trim().length > 0);
      const fallbackSource = firstTextField ? row[firstTextField.field] : "";
      row.sourceEn = sanitizeSheetCell(fallbackSource);
    }

    rows[i] = row;
  }

  return rows;
}

export function encodeTlkText(text: string): Uint8Array {
  return new TextEncoder().encode(String(text || ""));
}

export function quickChecksumHex(bytes: Uint8Array): string {
  let hash = 2166136261;
  for (let i = 0; i < bytes.length; i += 1) {
    hash ^= bytes[i];
    hash = Math.imul(hash, 16777619) >>> 0;
  }
  return hash.toString(16).toUpperCase().padStart(8, "0");
}

export function makeTlkFileName(baseName: string, locale: string, isDialogf: boolean): string {
  const suffix = isDialogf ? "dialogf" : "dialog";
  const token = localeCodeToFieldToken(locale).replace(/_/g, "-");
  return `${baseName}_${suffix}_${token}.tlk`;
}

export function buildSingleTlkBinaryFromColumn(rows: readonly TlkGridRow[], columnField: string, languageId: number): Uint8Array {
  const texts = rows.map(row => String((row as Record<string, unknown>)[columnField] || ""));
  const encodedRows = texts.map(encodeTlkText);
  const descriptorBytes = rows.length * TLK_ENTRY_SIZE;
  const stringBlobOffset = TLK_HEADER_SIZE + descriptorBytes;
  const stringBlobSize = encodedRows.reduce((sum, bytes) => sum + bytes.length, 0);
  const totalBytes = stringBlobOffset + stringBlobSize;
  const buffer = new ArrayBuffer(totalBytes);
  const view = new DataView(buffer);
  const bytes = new Uint8Array(buffer);

  bytes.set([0x54, 0x4c, 0x4b, 0x20], 0);
  bytes.set([0x56, 0x33, 0x2e, 0x30], 4);
  view.setUint32(8, Number(languageId) || 0, true);
  view.setUint32(12, rows.length, true);
  view.setUint32(16, stringBlobOffset, true);

  let textCursor = stringBlobOffset;
  for (let i = 0; i < rows.length; i += 1) {
    const descBase = TLK_HEADER_SIZE + i * TLK_ENTRY_SIZE;
    const textBytes = encodedRows[i];
    const hasText = textBytes.length > 0;
    const flags = hasText ? TLK_TEXT_PRESENT_FLAG : 0;
    view.setUint32(descBase, flags, true);
    view.setUint32(descBase + 28, hasText ? textCursor - stringBlobOffset : 0, true);
    view.setUint32(descBase + 32, textBytes.length, true);
    if (hasText) {
      bytes.set(textBytes, textCursor);
      textCursor += textBytes.length;
    }
  }

  return bytes;
}
