import type { LocaleColumn, TlkDiffReport, TlkGridRow } from "./types";

function tokenSignature(value: string): string {
  return String(value || "").match(/\{[A-Z0-9_]+\}/g)?.join("|") ?? "";
}

export interface DiffComputeInput {
  baselineRows: readonly TlkGridRow[];
  currentRows: readonly TlkGridRow[];
  localeColumns: readonly LocaleColumn[];
  previewLimit?: number;
}

export interface TlkGridRowWithDiffStatus {
  strRef: number;
  locale: string;
  before: string;
  after: string;
  statusLabel: string;
  statusClass: "status-ok" | "status-warn" | "status-error";
}

export interface DiffCounts {
  changed: number;
  added: number;
  removed: number;
  conflicts: number;
}

export function computeDiff(input: DiffComputeInput): { report: TlkDiffReport; rows: TlkGridRowWithDiffStatus[]; counts: DiffCounts } {
  const baselineMap = new Map<number, TlkGridRow>(input.baselineRows.map(row => [Number(row.strRef), row]));
  const currentMap = new Map<number, TlkGridRow>(input.currentRows.map(row => [Number(row.strRef), row]));
  const previewLimitValue = input.previewLimit;
  const previewLimit = Number.isFinite(previewLimitValue || 0) && (previewLimitValue || 0) > 0 ? Number(previewLimitValue) : 8;

  const allStrRefs = new Set<number>();
  input.baselineRows.forEach(row => allStrRefs.add(Number(row.strRef)));
  input.currentRows.forEach(row => allStrRefs.add(Number(row.strRef)));
  const sorted = Array.from(allStrRefs).sort((a, b) => a - b);

  const counts: DiffCounts = {
    changed: 0,
    added: 0,
    removed: 0,
    conflicts: 0
  };
  const rows: TlkGridRowWithDiffStatus[] = [];

  sorted.forEach(strRef => {
    const baseline = baselineMap.get(strRef);
    const current = currentMap.get(strRef);
    input.localeColumns.forEach(col => {
      const before = String((baseline as Record<string, unknown> | undefined)?.[col.field] || "");
      const after = String((current as Record<string, unknown> | undefined)?.[col.field] || "");
      if (before === after) return;

      let statusLabel = "Updated";
      let statusClass: "status-ok" | "status-warn" | "status-error" = "status-ok";

      counts.changed += 1;
      if (!before && after) {
        counts.added += 1;
        statusLabel = "Added";
      } else if (before && !after) {
        counts.removed += 1;
        statusLabel = "Removed in XLSX";
        statusClass = "status-warn";
      } else {
        const beforeTokens = tokenSignature(before);
        const afterTokens = tokenSignature(after);
        if (beforeTokens !== afterTokens) {
          counts.conflicts += 1;
          statusLabel = "Token mismatch";
          statusClass = "status-warn";
        }
      }

      if (rows.length < previewLimit) {
        rows.push({
          strRef,
          locale: col.title,
          before: before || "-",
          after: after || "-",
          statusLabel,
          statusClass
        });
      }
    });
  });

  const summary = {
    changed: counts.changed,
    added: counts.added,
    removed: counts.removed,
    conflicts: counts.conflicts
  };

  return {
    report: {
      generatedAt: new Date().toLocaleString(),
      summary,
      rows
    },
    rows,
    counts
  };
}

export function renderDiffMarkdown(report: TlkDiffReport): string {
  const lines = [
    "# TLK Diff Report",
    "",
    `Generated: ${report.generatedAt}`,
    "",
    `Changed: ${report.summary.changed}`,
    `Added: ${report.summary.added}`,
    `Removed: ${report.summary.removed}`,
    `Conflicts: ${report.summary.conflicts}`,
    "",
    "",
    "| StrRef | Locale | Before | After | Status |",
    "| --- | --- | --- | --- | --- |"
  ];

  report.rows.forEach(item => {
    lines.push(`| ${item.strRef} | ${item.locale} | ${item.before} | ${item.after} | ${item.statusLabel} |`);
  });

  return lines.join("\n");
}
