/// <reference lib="webworker" />

import { buildXlsxHeaderForExport, sanitizeSheetCell } from "../lib/xlsx";
import type { LocaleColumn, TlkGridRow } from "../lib/types";
import cptable from "codepage";

type ExportRequest = {
  rows: TlkGridRow[];
  localeColumns: LocaleColumn[];
  fileName: string;
  chunkSize?: number;
  lineEnding?: "lf" | "crlf";
  encoding?:
    | "utf8"
    | "utf8-bom"
    | "iso-8859-1"
    | "windows-1252"
    | "windows-1251"
    | "euc-jp"
    | "shift_jis"
    | "euc-kr"
    | "windows-1250"
    | "gb2312"
    | "iso-8859-2";
};

type ProgressMessage = {
  type: "progress";
  doneRows: number;
  totalRows: number;
  progress: number;
};

type ResultMessage = {
  type: "result";
  fileName: string;
  bytes: ArrayBuffer;
};

type ErrorMessage = {
  type: "error";
  error: string;
};

const ENCODING_TO_CODEPAGE: Record<string, number> = {
  "iso-8859-1": 28591,
  "windows-1252": 1252,
  "windows-1251": 1251,
  "euc-jp": 51932,
  shift_jis: 932,
  "euc-kr": 51949,
  "windows-1250": 1250,
  gb2312: 936,
  "iso-8859-2": 28592,
};

function escapeCsvCell(value: string | number): string {
  const raw = String(value);
  if (raw.includes('"')) {
    const quoted = raw.replace(/"/g, '""');
    return `"${quoted}"`;
  }
  if (raw.includes(",") || raw.includes("\n") || raw.includes("\r")) {
    return `"${raw}"`;
  }
  return raw;
}

function encodeCsvContent(content: string, encoding: ExportRequest["encoding"]): Uint8Array {
  if (!encoding || encoding === "utf8" || encoding === "utf8-bom") {
    const bodyBytes = new TextEncoder().encode(content);
    if (encoding === "utf8-bom") {
      const withBom = new Uint8Array(bodyBytes.byteLength + 3);
      withBom.set([0xef, 0xbb, 0xbf], 0);
      withBom.set(bodyBytes, 3);
      return withBom;
    }
    return bodyBytes;
  }

  const codepage = ENCODING_TO_CODEPAGE[encoding];
  if (!codepage) {
    throw new Error(`Unsupported CSV encoding: ${encoding}`);
  }

  const encoded = cptable.utils.encode(codepage, content, "arr");
  if (Array.isArray(encoded)) {
    return Uint8Array.from(encoded);
  }
  if (encoded instanceof Uint8Array) {
    return encoded;
  }
  if (typeof encoded === "string") {
    const bytes = new Uint8Array(encoded.length);
    for (let i = 0; i < encoded.length; i += 1) {
      bytes[i] = encoded.charCodeAt(i) & 0xff;
    }
    return bytes;
  }
  throw new Error(`Failed to encode CSV content for ${encoding}.`);
}

const workerScope: DedicatedWorkerGlobalScope = self as unknown as DedicatedWorkerGlobalScope;

workerScope.onmessage = (event: MessageEvent<ExportRequest>) => {
  try {
    const { rows, localeColumns, fileName, chunkSize, lineEnding, encoding } = event.data;
    const totalRows = rows.length;
    const step = Math.max(1000, Number(chunkSize || 5000));
    const headers = buildXlsxHeaderForExport(localeColumns);

    const lines: string[] = [];
    lines.push(headers.map(escapeCsvCell).join(","));

    for (let offset = 0; offset < totalRows; offset += step) {
      const end = Math.min(totalRows, offset + step);
      for (let index = offset; index < end; index += 1) {
        const row = rows[index];
        const csvRow: Array<string | number> = [row.strRef];

        for (let colIndex = 0; colIndex < localeColumns.length; colIndex += 1) {
          const col = localeColumns[colIndex];
          csvRow.push(sanitizeSheetCell(row[col.field]));
        }
        lines.push(csvRow.map(escapeCsvCell).join(","));
      }

      const doneRows = end;
      const progress = totalRows > 0 ? Math.round((doneRows / totalRows) * 100) : 100;
      const progressMessage: ProgressMessage = { type: "progress", doneRows, totalRows, progress };
      workerScope.postMessage(progressMessage);
    }

    const eol = lineEnding === "crlf" ? "\r\n" : "\n";
    const content = lines.join(eol);
    const bytes = encodeCsvContent(content, encoding);
    const cloned = bytes.slice();
    const resultMessage: ResultMessage = { type: "result", fileName, bytes: cloned.buffer };
    workerScope.postMessage(resultMessage, [cloned.buffer]);
  } catch (error) {
    const message = error instanceof Error ? error.message : "CSV export worker failed.";
    const errorMessage: ErrorMessage = { type: "error", error: message };
    workerScope.postMessage(errorMessage);
  }
};
