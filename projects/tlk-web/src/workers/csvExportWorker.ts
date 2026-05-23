/// <reference lib="webworker" />

import { buildXlsxHeaderForExport, sanitizeSheetCell } from "../lib/xlsx";
import type { LocaleColumn, TlkGridRow } from "../lib/types";

type ExportRequest = {
  rows: TlkGridRow[];
  localeColumns: LocaleColumn[];
  fileName: string;
  chunkSize?: number;
  lineEnding?: "lf" | "crlf";
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

const workerScope: DedicatedWorkerGlobalScope = self as unknown as DedicatedWorkerGlobalScope;

workerScope.onmessage = (event: MessageEvent<ExportRequest>) => {
  try {
    const { rows, localeColumns, fileName, chunkSize, lineEnding } = event.data;
    const totalRows = rows.length;
    const step = Math.max(1000, Number(chunkSize || 5000));
    const headers = buildXlsxHeaderForExport(localeColumns);

    const lines: string[] = [];
    lines.push(headers.map(escapeCsvCell).join(","));

    for (let offset = 0; offset < totalRows; offset += step) {
      const end = Math.min(totalRows, offset + step);
      for (let index = offset; index < end; index += 1) {
        const row = rows[index];
        const csvRow: Array<string | number> = [
          row.strRef,
          sanitizeSheetCell(row.sourceEn),
          sanitizeSheetCell(row.status || "Draft"),
        ];

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
    const bytes = new TextEncoder().encode(content);
    const cloned = bytes.slice();
    const resultMessage: ResultMessage = { type: "result", fileName, bytes: cloned.buffer };
    workerScope.postMessage(resultMessage, [cloned.buffer]);
  } catch (error) {
    const message = error instanceof Error ? error.message : "CSV export worker failed.";
    const errorMessage: ErrorMessage = { type: "error", error: message };
    workerScope.postMessage(errorMessage);
  }
};
