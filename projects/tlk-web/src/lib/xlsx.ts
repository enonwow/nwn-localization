import type { LocaleColumn, LocaleVariant, TlkGridRow, ValidationIssue, ValidationResult, XlsxRow } from "./types";
import { localeCodeToFieldToken, normalizeLocaleCode } from "./types";

export interface XlsxRuntime {
  read(data: ArrayBuffer, options: { type: "array" }): {
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
    const status = sanitizeSheetCell(row.Status);
    const result: TlkGridRow = {
      id: Number(row.StrRef),
      strRef: Number(row.StrRef),
      sourceEn: sanitizeSheetCell(row.Source_EN || row.Source || ""),
      context: sanitizeSheetCell(row.Context || ""),
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
      Source_EN: row.sourceEn,
      Status: row.status || "Draft"
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
  const headers = ["StrRef", "Source_EN", "Status"];
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
      sanitizeSheetCell(row.sourceEn),
      sanitizeSheetCell(row.status || "Draft"),
    ];

    for (let colIndex = 0; colIndex < localeColumns.length; colIndex += 1) {
      const col = localeColumns[colIndex];
      line.push(sanitizeSheetCell(row[col.field]));
    }
    result.push(line);
  }

  return result;
}

export function makeXlsxFileName(): string {
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `localization_sheet_v4_${timestamp}.xlsx`;
}

export function makeCsvFileName(): string {
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `localization_sheet_v4_${timestamp}.csv`;
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
  const workbook = xlsx.read(buffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("XLSX workbook has no sheets.");
  }
  const sheet = workbook.Sheets[firstSheetName];
  const jsonRows = xlsx.utils.sheet_to_json(sheet, {
    defval: "",
    raw: false,
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
  const hasSource = firstRow.Source_EN !== undefined || firstRow.Source !== undefined;
  const hasStrRef = firstRow.StrRef !== undefined;
  if (!hasStrRef) {
    issues.push({ severity: "error", field: "StrRef", message: `${sourceLabel}: missing required column StrRef.` });
  }
  if (!hasSource) {
    issues.push({ severity: "error", message: `${sourceLabel}: missing required column Source_EN (or Source).` });
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
