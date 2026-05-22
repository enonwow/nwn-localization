import { getParsedLocaleColumnsFromRows, validateXlsxSchema, type ParsedXlsxWorkbook } from "./xlsx";
import {
  type BundleFallbackMeta,
  type DialogfFallbackResult,
  type TlkBundleConfig,
  type ValidationIssue,
  type ValidationResult,
  type ValidationSeverity,
  type TlkGridRow,
  normalizeLocaleCode
} from "./types";

interface ValidationState extends ValidationResult {
  normalized: TlkBundleConfig[];
}

function duplicateLocaleList(values: readonly string[]): string[] {
  const counter = new Map<string, number>();
  values.forEach(value => {
    if (!value) return;
    counter.set(value, (counter.get(value) || 0) + 1);
  });

  return Array.from(counter.entries())
    .filter(([, count]) => count > 1)
    .map(([value]) => value);
}

function isTlkFile(value: string): boolean {
  return value.toLowerCase().endsWith(".tlk");
}

export function validateTlkBundles(configs: readonly TlkBundleConfig[]): ValidationState {
  const issues: ValidationIssue[] = [];

  if (!Array.isArray(configs) || configs.length === 0) {
    return {
      ok: false,
      normalized: [],
      issues: [{ severity: "error", message: "Add at least one TLK locale pack." }]
    };
  }

  const normalized = configs.map((bundle, index) => {
    const locale = normalizeLocaleCode(bundle.locale);
    const dialog = String(bundle.dialog || "").trim();
    const dialogf = String(bundle.dialogf || "").trim();
    const dialogfAuto = Boolean(bundle.dialogfAuto);
    const normalizedBundle: TlkBundleConfig = {
      id: bundle.id,
      locale,
      dialog,
      dialogf,
      dialogfAuto
    };

    if (!locale) {
      issues.push({
        severity: "error",
        message: `Row ${index + 1}: locale is required.`,
        field: `locale:${index}`
      });
    }
    if (!dialog) {
      issues.push({
        severity: "error",
        message: `${rowLabel(index)}: dialog.tlk is required.`,
        field: `dialog:${index}`
      });
    } else if (!isTlkFile(dialog)) {
      issues.push({
        severity: "error",
        message: `${rowLabel(index)}: dialog file must end with .tlk.`,
        field: `dialog:${index}`
      });
    }
    if (dialogf && !isTlkFile(dialogf)) {
      issues.push({
        severity: "error",
        message: `${rowLabel(index)}: dialogf file must end with .tlk or stay empty for fallback.`,
        field: `dialogf:${index}`
      });
    }

    return normalizedBundle;
  });

  const duplicates = duplicateLocaleList(normalized.map(bundle => bundle.locale));
  if (duplicates.length > 0) {
    normalized.forEach(bundle => {
      if (duplicates.includes(bundle.locale)) {
        issues.push({
          severity: "error",
          message: `Duplicate locale: ${bundle.locale}.`,
          field: `locale:${bundle.locale}`
        });
      }
    });
  }

  return {
    ok: !issues.some(issue => issue.severity === "error"),
    normalized,
    issues
  };
}

export function applyDialogfFallbacks(configs: readonly TlkBundleConfig[]): DialogfFallbackResult {
  const metadata: BundleFallbackMeta[] = [];
  const fallbackBundles = configs.map(bundle => {
    const next = {
      id: bundle.id,
      locale: normalizeLocaleCode(bundle.locale),
      dialog: String(bundle.dialog || "").trim(),
      dialogf: String(bundle.dialogf || "").trim(),
      dialogfAuto: Boolean(bundle.dialogfAuto)
    };

    if (next.dialog && !next.dialogf) {
      metadata.push({
        locale: next.locale,
        dialog: next.dialog,
        source: "dialogfAutoFallback"
      });
      next.dialogf = next.dialog;
      next.dialogfAuto = true;
    }
    return next;
  });

  return {
    bundles: fallbackBundles,
    metadata,
    fallbackCount: metadata.length
  };
}

function rowLabel(index: number): string {
  return `Row ${index + 1}`;
}

export function validateSchemaFromXlsxWorkbook(
  workbook: ParsedXlsxWorkbook,
  sourceRows: readonly unknown[]
): ValidationResult {
  const normalizedRows = sourceRows as readonly TlkGridRow[];
  const result = validateXlsxSchema(
    normalizedRows.map(row => ({ ...row })),
    `XLSX source (${workbook.sheetName})`
  );
  const parsedColumns = getParsedLocaleColumnsFromRows(normalizedRows);
  const extraWarnings: ValidationIssue[] = [];

  if (parsedColumns.length === 0) {
    extraWarnings.push({
      severity: "warning",
      message: "No locale columns found in normalized rows."
    });
  }

  return {
    ok: result.ok,
    issues: [...result.issues, ...extraWarnings]
  };
}

export function combineValidationResult(...results: ValidationResult[]): ValidationResult {
  const issues = results.flatMap(result => result.issues);
  return {
    ok: issues.every(issue => issue.severity !== "error"),
    issues
  };
}

export function isValidationWarning(severity: ValidationSeverity): boolean {
  return severity === "warning";
}
