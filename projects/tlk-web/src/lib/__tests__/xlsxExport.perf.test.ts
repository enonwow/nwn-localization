import { describe, expect, it } from "vitest";

import { buildXlsxAoaForExport } from "../xlsx";
import type { LocaleColumn, TlkGridRow } from "../types";

function makeRows(count: number): TlkGridRow[] {
  const rows: TlkGridRow[] = [];
  for (let i = 0; i < count; i += 1) {
    rows.push({
      id: i,
      strRef: i,
      sourceEn: `Source ${i}`,
      context: "",
      status: "Draft",
      loc_en: `EN ${i}`,
      loc_pl: `PL ${i}`,
      loc_de: `DE ${i}`,
      loc_fr: i % 4 === 0 ? "" : `FR ${i}`,
    });
  }
  return rows;
}

describe("xlsx export performance smoke", () => {
  it("builds AOA payload for large datasets within budget", () => {
    const localeColumns: LocaleColumn[] = [
      { field: "loc_en", title: "EN", locale: "EN", variant: "dialog" },
      { field: "loc_pl", title: "PL", locale: "PL", variant: "dialog" },
      { field: "loc_de", title: "DE", locale: "DE", variant: "dialog" },
      { field: "loc_fr", title: "FR", locale: "FR", variant: "dialog" },
    ];
    const rowCount = Number(process.env.PERF_ROW_COUNT || "20000");
    const maxMs = Number(process.env.PERF_EXPORT_MAX_MS || "5000");
    const rows = makeRows(rowCount);

    const start = Date.now();
    const aoa = buildXlsxAoaForExport(rows, localeColumns);
    const elapsedMs = Date.now() - start;

    expect(aoa.length).toBe(rowCount + 1);
    expect(aoa[0].length).toBe(7);
    expect(elapsedMs).toBeLessThan(maxMs);
  });
});

