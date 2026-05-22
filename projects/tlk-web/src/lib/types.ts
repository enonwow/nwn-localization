export type LocaleVariant = "dialog" | "dialogf" | "xlsx-extra";

export interface LocaleColumn {
  field: string;
  title: string;
  locale: string;
  variant: LocaleVariant;
}

export interface TlkBundleConfig {
  id?: number | string;
  locale: string;
  dialog: string;
  dialogf: string;
  dialogfAuto: boolean;
}

export interface ParsedTlk {
  fileName: string;
  languageId: number;
  entryCount: number;
  rows: string[];
}

export interface ParsedTlkBundle {
  bundle: TlkBundleConfig;
  parsed: ParsedTlk;
  dialogfParsed: ParsedTlk | null;
}

export type ValidationSeverity = "error" | "warning" | "success";

export interface ValidationIssue {
  severity: ValidationSeverity;
  message: string;
  field?: string;
}

export interface ValidationResult {
  ok: boolean;
  issues: ValidationIssue[];
}

export interface TlkGridRow {
  id: number;
  strRef: number;
  sourceEn: string;
  context: string;
  status: string;
  [key: string]: string | number | undefined;
}

export type XlsxRow = Record<string, string | number | boolean | null | undefined>;

export interface TlkDiffRow {
  strRef: number;
  locale: string;
  before: string;
  after: string;
  statusLabel: string;
  statusClass: "status-ok" | "status-warn" | "status-error";
}

export interface TlkDiffSummary {
  changed: number;
  added: number;
  removed: number;
  conflicts: number;
}

export interface TlkDiffReport {
  generatedAt: string;
  summary: TlkDiffSummary;
  rows: TlkDiffRow[];
}

export interface BundleFallbackMeta {
  locale: string;
  dialog: string;
  source: "dialogfAutoFallback";
}

export interface DialogfFallbackResult {
  bundles: TlkBundleConfig[];
  metadata: BundleFallbackMeta[];
  fallbackCount: number;
}

export interface TlkSourceXlsxMeta {
  kind: "xlsx";
  fileName: string;
  sheetName: string;
}

export interface TlkSourceTlkMeta {
  kind: "tlk";
  entryCount: number;
  fallbackCount: number;
}

export type TlkSourceMeta = TlkSourceXlsxMeta | TlkSourceTlkMeta;

export function normalizeLocaleCode(value: string): string {
  return String(value || "").toUpperCase().trim();
}

export function localeCodeToFieldToken(localeCode: string): string {
  return String(localeCode || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}
