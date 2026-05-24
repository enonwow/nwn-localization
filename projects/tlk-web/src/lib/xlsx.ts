import type { LocaleColumn, LocaleVariant, TlkGridRow, ValidationIssue, ValidationResult, XlsxRow } from "./types";
import { localeCodeToFieldToken, normalizeLocaleCode } from "./types";

export interface XlsxRuntime {
  read(data: ArrayBuffer, options: { type: "array"; raw?: boolean; cellDates?: boolean }): {
    SheetNames: string[];
    Sheets: Record<string, unknown>;
  };
  utils: {
    sheet_to_json(
      sheet: unknown,
      options: {
        defval: string;
        raw: boolean;
        blankrows: boolean;
      }
    ): unknown[];
  };
}

export interface XlsxFile {
  arrayBuffer(): Promise<ArrayBuffer>;
  name: string;
}

export interface ParsedXlsxWorkbook {
  workbook: unknown;
  sheetName: string;
  rows: XlsxRow[];
}

export function sanitizeSheetCell(value: unknown): string {
  if (value == null) return "";
  return String(value).replace(/\r\n/g, "\n");
}

export function localeColumnsFromXlsxColumns(columns: readonly string[]): LocaleColumn[] {
  return columns
    .filter(columnName => /^loc_[a-z0-9_]+/i.test(columnName))
    .map(colName => {
      const suffix = colName.slice(4);
      const isDialogf = suffix.endsWith("_f");
      const isExtra = suffix.endsWith("_xlsx");
      const localeToken = isDialogf
        ? suffix.slice(0, -2)
        : isExtra
          ? suffix.slice(0, -5)
          : suffix;
      const locale = localeToken.replace(/_/g, "-").toUpperCase();
      return {
        field: colName,
        title: isDialogf ? `${locale} F` : isExtra ? `${locale} XLS` : locale,
        locale,
        variant: (isDialogf ? "dialogf" : isExtra ? "xlsx-extra" : "dialog") as LocaleVariant
      };
    });
}

export function buildTlkRowsFromXlsxRows(sheetRows: readonly XlsxRow[]): TlkGridRow[] {
  const normalized = sheetRows
    .map(row => row || {})
    .filter(row => row.StrRef !== undefined && row.StrRef !== null && String(row.StrRef).trim() !== "")
    .filter(row => Number.isFinite(Number(row.StrRef)));

  normalized.sort((a, b) => Number(a.StrRef) - Number(b.StrRef));

  return normalized.map(row => {
    const localeKeyOrder = Object.keys(row).filter((key) => {
      if (["StrRef", "Source_EN", "Source", "Context", "Status"].includes(key)) return false;
      return true;
    });
    const firstLocaleKey = localeKeyOrder[0];
    const fallbackSource = firstLocaleKey ? sanitizeSheetCell(row[firstLocaleKey]) : "";
    const status = sanitizeSheetCell(row.Status);
    const result: TlkGridRow = {
      id: Number(row.StrRef),
      strRef: Number(row.StrRef),
      sourceEn: sanitizeSheetCell(row.Source_EN || row.Source || fallbackSource),
      context: "",
      status: status || "Draft"
    };

    Object.keys(row).reduce<TlkGridRow>((acc, key) => {
      if (["StrRef", "Source_EN", "Source", "Context", "Status"].includes(key)) {
        return acc;
      }

      const localeLabel = String(key).trim();
      const lower = localeLabel.toLowerCase();
      let field = "";
      if (lower.endsWith("_f")) {
        field = `loc_${localeCodeToFieldToken(localeLabel.slice(0, -2))}_f`;
      } else if (lower.endsWith("_xls")) {
        field = `loc_${localeCodeToFieldToken(localeLabel.slice(0, -4))}_xlsx`;
      } else {
        field = `loc_${localeCodeToFieldToken(localeLabel)}`;
      }
      acc[field] = sanitizeSheetCell(row[key]);
      return acc;
    }, result);

    return result;
  });
}

export function buildXlsxRowsForExport(rows: readonly TlkGridRow[], localeColumns: readonly LocaleColumn[]): XlsxRow[] {
  return rows.map(row => {
    const exported: XlsxRow = {
      StrRef: row.strRef,
    };

    localeColumns.forEach(col => {
      const columnName = col.variant === "dialogf"
        ? `${col.locale}_F`
        : col.variant === "xlsx-extra"
          ? `${col.locale}_XLS`
          : col.locale;
      exported[columnName] = sanitizeSheetCell(row[col.field]);
    });

    return exported;
  });
}

export function buildXlsxHeaderForExport(localeColumns: readonly LocaleColumn[]): string[] {
  const headers = ["StrRef"];
  for (let i = 0; i < localeColumns.length; i += 1) {
    const col = localeColumns[i];
    const columnName = col.variant === "dialogf"
      ? `${col.locale}_F`
      : col.variant === "xlsx-extra"
        ? `${col.locale}_XLS`
        : col.locale;
    headers.push(columnName);
  }
  return headers;
}

export function buildXlsxAoaForExport(
  rows: readonly TlkGridRow[],
  localeColumns: readonly LocaleColumn[],
): Array<Array<string | number>> {
  const headers = buildXlsxHeaderForExport(localeColumns);
  const result: Array<Array<string | number>> = [headers];

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex];
    const line: Array<string | number> = [
      row.strRef,
    ];

    for (let colIndex = 0; colIndex < localeColumns.length; colIndex += 1) {
      const col = localeColumns[colIndex];
      line.push(sanitizeSheetCell(row[col.field]));
    }
    result.push(line);
  }

  return result;
}

export interface CsvExportChangeInput {
  baselineRows: readonly TlkGridRow[];
  currentRows: readonly TlkGridRow[];
  localeColumns: readonly LocaleColumn[];
}

export interface CsvExportChangeSummary {
  changedRows: number;
  changedCells: number;
  addedRows: number;
  removedRows: number;
  sourceChangedRows: number;
  changedRowStrRefs: number[];
  touchedLocales: string[];
}

export function computeCsvExportChangeSummary(input: CsvExportChangeInput): CsvExportChangeSummary {
  const baselineByStrRef = new Map<number, TlkGridRow>();
  const currentByStrRef = new Map<number, TlkGridRow>();
  const touchedLocales = new Set<string>();
  const changedRowStrRefs: number[] = [];

  for (let i = 0; i < input.baselineRows.length; i += 1) {
    const row = input.baselineRows[i];
    baselineByStrRef.set(Number(row.strRef), row);
  }

  for (let i = 0; i < input.currentRows.length; i += 1) {
    const row = input.currentRows[i];
    currentByStrRef.set(Number(row.strRef), row);
  }

  let changedRows = 0;
  let changedCells = 0;
  let addedRows = 0;
  let removedRows = 0;
  let sourceChangedRows = 0;

  const allStrRefs = new Set<number>([
    ...baselineByStrRef.keys(),
    ...currentByStrRef.keys(),
  ]);

  allStrRefs.forEach((strRef) => {
    const baselineRow = baselineByStrRef.get(strRef);
    const currentRow = currentByStrRef.get(strRef);

    if (!baselineRow && currentRow) {
      addedRows += 1;
      changedRows += 1;
      changedRowStrRefs.push(strRef);
      const sourceValue = sanitizeSheetCell(currentRow.sourceEn);
      if (sourceValue.length > 0) {
        sourceChangedRows += 1;
        changedCells += 1;
      }
      for (let i = 0; i < input.localeColumns.length; i += 1) {
        const col = input.localeColumns[i];
        const value = sanitizeSheetCell(currentRow[col.field]);
        if (value.length > 0) {
          touchedLocales.add(col.title);
          changedCells += 1;
        }
      }
      return;
    }

    if (baselineRow && !currentRow) {
      removedRows += 1;
      changedRows += 1;
      changedRowStrRefs.push(strRef);
      const sourceValue = sanitizeSheetCell(baselineRow.sourceEn);
      if (sourceValue.length > 0) {
        sourceChangedRows += 1;
        changedCells += 1;
      }
      for (let i = 0; i < input.localeColumns.length; i += 1) {
        const col = input.localeColumns[i];
        const value = sanitizeSheetCell(baselineRow[col.field]);
        if (value.length > 0) {
          touchedLocales.add(col.title);
          changedCells += 1;
        }
      }
      return;
    }

    if (!baselineRow || !currentRow) {
      return;
    }

    let rowChanged = false;
    const baselineSource = sanitizeSheetCell(baselineRow.sourceEn);
    const currentSource = sanitizeSheetCell(currentRow.sourceEn);
    if (baselineSource !== currentSource) {
      rowChanged = true;
      sourceChangedRows += 1;
      changedCells += 1;
    }

    for (let i = 0; i < input.localeColumns.length; i += 1) {
      const col = input.localeColumns[i];
      const currentHasField = Object.prototype.hasOwnProperty.call(currentRow, col.field);
      const baselineHasField = Object.prototype.hasOwnProperty.call(baselineRow, col.field);
      if (currentHasField !== baselineHasField) {
        rowChanged = true;
        touchedLocales.add(col.title);
        changedCells += 1;
        continue;
      }

      const baselineValue = sanitizeSheetCell(baselineRow[col.field]);
      const currentValue = sanitizeSheetCell(currentRow[col.field]);
      if (baselineValue !== currentValue) {
        rowChanged = true;
        touchedLocales.add(col.title);
        changedCells += 1;
      }
    }

    if (rowChanged) {
      changedRows += 1;
      changedRowStrRefs.push(strRef);
    }
  });

  changedRowStrRefs.sort((a, b) => a - b);

  return {
    changedRows,
    changedCells,
    addedRows,
    removedRows,
    sourceChangedRows,
    changedRowStrRefs,
    touchedLocales: Array.from(touchedLocales).sort((a, b) => a.localeCompare(b)),
  };
}

export function hasCsvExportChanges(input: CsvExportChangeInput): boolean {
  return computeCsvExportChangeSummary(input).changedRows > 0;
}

export function makeXlsxFileName(): string {
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `localization_sheet_v4_${timestamp}.xlsx`;
}

function basenameFromPath(value: string): string {
  const normalized = value.replace(/\\/g, "/");
  const lastSlash = normalized.lastIndexOf("/");
  return lastSlash >= 0 ? normalized.slice(lastSlash + 1) : normalized;
}

export function makeCsvFileName(sourceName?: string): string {
  const fallback = "latest-localization.csv";
  const raw = String(sourceName || "").trim();
  if (!raw) return fallback;

  const fileName = basenameFromPath(raw);
  const stem = fileName.replace(/\.[^.]+$/, "").trim();
  if (!stem) return fallback;

  const safeStem = stem.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_");
  if (!safeStem) return fallback;
  return `${safeStem}.csv`;
}

export function hasXlsxRuntime(xlsx: XlsxRuntime | undefined): xlsx is XlsxRuntime {
  return !!xlsx && typeof xlsx.read === "function" && typeof xlsx.utils?.sheet_to_json === "function";
}

export async function parseWorkbookFile(file: XlsxFile, xlsx: XlsxRuntime): Promise<ParsedXlsxWorkbook> {
  if (!file) {
    throw new Error("Select XLSX file first.");
  }
  if (!hasXlsxRuntime(xlsx)) {
    throw new Error("XLSX runtime failed to load.");
  }

  const buffer = await file.arrayBuffer();
  const workbook = xlsx.read(buffer, {
    type: "array",
    raw: true,
    cellDates: false,
  });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("XLSX workbook has no sheets.");
  }
  const sheet = workbook.Sheets[firstSheetName];
  const jsonRows = xlsx.utils.sheet_to_json(sheet, {
    defval: "",
    raw: true,
    blankrows: false
  });
  if (!Array.isArray(jsonRows) || jsonRows.length === 0) {
    throw new Error("XLSX has no data rows.");
  }

  return {
    workbook,
    sheetName: firstSheetName,
    rows: jsonRows as XlsxRow[]
  };
}

export function validateXlsxSchema(rows: readonly XlsxRow[], sourceLabel = "XLSX source"): ValidationResult {
  const issues: ValidationIssue[] = [];
  if (!Array.isArray(rows) || rows.length === 0) {
    return {
      ok: false,
      issues: [{ severity: "error", message: `${sourceLabel}: has no rows.` }]
    };
  }

  const firstRow = rows[0] as Record<string, unknown>;
  const hasStrRef = firstRow.StrRef !== undefined;
  if (!hasStrRef) {
    issues.push({ severity: "error", field: "StrRef", message: `${sourceLabel}: missing required column StrRef.` });
  }
  if (!Object.keys(firstRow).some(key => /^loc_/i.test(key))) {
    issues.push({ severity: "warning", message: `${sourceLabel}: no locale columns detected; export/import may be translation-only.` });
  }

  const badStrRefCount = rows.filter(row => Number.isNaN(Number((row as Record<string, unknown>).StrRef))).length;
  if (badStrRefCount > 0) {
    issues.push({ severity: "error", message: `${sourceLabel}: found ${badStrRefCount} row(s) with invalid StrRef.` });
  }

  const locales = Object.keys(firstRow)
    .filter(key => /^loc_/i.test(key))
    .map(key => normalizeLocaleCode(key.slice(4)));
  const uniqueLocales = new Set(locales);
  if (uniqueLocales.size !== locales.length) {
    issues.push({ severity: "warning", message: `${sourceLabel}: duplicate locale columns detected in first row.` });
  }

  return {
    ok: issues.every(issue => issue.severity !== "error"),
    issues
  };
}

export function getParsedLocaleColumnsFromRows(rows: readonly TlkGridRow[]): LocaleColumn[] {
  if (!rows.length) return [];
  const firstRow = rows[0];
  return Object.keys(firstRow)
    .filter(field => /^loc_[a-z0-9_]+/i.test(field))
    .map(field => {
      const suffix = field.slice(4);
      const isDialogf = suffix.endsWith("_f");
      const isExtra = suffix.endsWith("_xlsx");
      const localeToken = isDialogf ? suffix.slice(0, -2) : isExtra ? suffix.slice(0, -5) : suffix;
      const locale = localeToken.replace(/_/g, "-").toUpperCase();
      return {
        field,
        title: isDialogf ? `${locale} F` : isExtra ? `${locale} XLS` : locale,
        locale,
        variant: (isDialogf ? "dialogf" : isExtra ? "xlsx-extra" : "dialog") as LocaleVariant
      };
    });
}
