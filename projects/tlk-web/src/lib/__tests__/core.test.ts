import { describe, expect, it, vi } from "vitest";

import { computeDiff, renderDiffMarkdown } from "../diff";
import { fetchRepoFileFromGitHub, publishCsvToGitHub } from "../github";
import {
  buildRowsFromParsedTlkBundles,
  buildSingleTlkBinaryFromColumn,
  fileNameToken,
  makeTlkFileName,
  parseSingleTlkBuffer,
  quickChecksumHex,
  readAscii,
  safeFileNameFromPath,
  sanitizeSheetCell as sanitizeTlkCell,
  TLK_LOCALE_TO_LANGUAGE_ID,
} from "../tlk";
import type { LocaleColumn, ParsedTlkBundle, TlkBundleConfig, TlkGridRow, ValidationResult, XlsxRow } from "../types";
import {
  applyDialogfFallbacks,
  combineValidationResult,
  isValidationWarning,
  validateSchemaFromXlsxWorkbook,
  validateTlkBundles,
} from "../validation";
import {
  buildXlsxAoaForExport,
  buildXlsxHeaderForExport,
  buildTlkRowsFromXlsxRows,
  computeCsvExportChangeSummary,
  buildXlsxRowsForExport,
  getParsedLocaleColumnsFromRows,
  hasCsvExportChanges,
  hasXlsxRuntime,
  localeColumnsFromXlsxColumns,
  makeCsvFileName,
  makeXlsxFileName,
  parseWorkbookFile,
  sanitizeSheetCell as sanitizeXlsxCell,
  validateXlsxSchema,
} from "../xlsx";

function makeParsedBundle(bundle: TlkBundleConfig, rows: string[], dialogfRows: string[] | null = null): ParsedTlkBundle {
  return {
    bundle,
    parsed: {
      fileName: bundle.dialog,
      languageId: TLK_LOCALE_TO_LANGUAGE_ID[bundle.locale] ?? 0,
      entryCount: rows.length,
      rows,
    },
    dialogfParsed: dialogfRows
      ? {
          fileName: bundle.dialogf || "dialogf.tlk",
          languageId: TLK_LOCALE_TO_LANGUAGE_ID[bundle.locale] ?? 0,
          entryCount: dialogfRows.length,
          rows: dialogfRows,
        }
      : null,
  };
}

function makeGridRows(): TlkGridRow[] {
  return [
    { id: 0, strRef: 0, sourceEn: "Hello", context: "UI", status: "Draft", loc_en: "Hello", loc_pl: "Czesc" },
    { id: 1, strRef: 1, sourceEn: "World", context: "World", status: "Draft", loc_en: "World", loc_pl: "Swiat" },
  ];
}

describe("TLK core", () => {
  it("parses TLK V3.0 buffer built from rows", () => {
    const rows = makeGridRows();
    const bytes = buildSingleTlkBinaryFromColumn(rows, "loc_en", 0);
    const copy = new Uint8Array(bytes.byteLength);
    copy.set(bytes);
    const parsed = parseSingleTlkBuffer(copy.buffer, "dialog_en.tlk");

    expect(parsed.entryCount).toBe(2);
    expect(parsed.rows[0]).toBe("Hello");
    expect(parsed.rows[1]).toBe("World");
  });

  it("handles empty TLK text entries", () => {
    const rows: TlkGridRow[] = [
      { id: 0, strRef: 0, sourceEn: "", context: "", status: "Draft", loc_en: "" },
      { id: 1, strRef: 1, sourceEn: "X", context: "", status: "Draft", loc_en: "X" },
    ];
    const bytes = buildSingleTlkBinaryFromColumn(rows, "loc_en", 0);
    const parsed = parseSingleTlkBuffer(bytes, "dialog_en.tlk");
    expect(parsed.rows[0]).toBe("");
    expect(parsed.rows[1]).toBe("X");
  });

  it("throws when TLK buffer is too small", () => {
    expect(() => parseSingleTlkBuffer(new ArrayBuffer(2), "bad.tlk")).toThrow("too small");
  });

  it("throws when TLK signature/version is invalid", () => {
    const raw = new Uint8Array(20);
    raw.set([0x42, 0x41, 0x44, 0x20], 0);
    raw.set([0x56, 0x31, 0x2e, 0x30], 4);
    expect(() => parseSingleTlkBuffer(raw.buffer, "bad.tlk")).toThrow("invalid TLK signature/version");
  });

  it("throws when TLK descriptor layout is malformed", () => {
    const raw = new Uint8Array(20);
    raw.set([0x54, 0x4c, 0x4b, 0x20], 0); // TLK
    raw.set([0x56, 0x33, 0x2e, 0x30], 4); // V3.0
    const view = new DataView(raw.buffer);
    view.setUint32(8, 0, true); // language
    view.setUint32(12, 1, true); // entryCount
    view.setUint32(16, 999, true); // invalid stringBlobOffset
    expect(() => parseSingleTlkBuffer(raw.buffer, "broken.tlk")).toThrow("malformed TLK layout");
  });

  it("provides utility helpers", () => {
    const bytes = new Uint8Array([0x54, 0x4c, 0x4b, 0x20]);
    const text = readAscii(new DataView(bytes.buffer), 0, 4);
    expect(text).toBe("TLK ");
    expect(fileNameToken("  Dialog_EN.TLK ")).toBe("dialog_en.tlk");
    expect(safeFileNameFromPath("C:\\games\\nwn\\dialog.tlk")).toBe("dialog.tlk");
    expect(sanitizeTlkCell("A\r\nB")).toBe("A\nB");
    expect(makeTlkFileName("rebuilt", "PT-BR", true)).toBe("rebuilt_dialogf_pt-br.tlk");
    expect(quickChecksumHex(new Uint8Array([1, 2, 3]))).toHaveLength(8);
  });

  it("builds grid rows from parsed bundles including dialogf column", () => {
    const bundles: ParsedTlkBundle[] = [
      makeParsedBundle(
        { id: 1, locale: "EN", dialog: "dialog_en.tlk", dialogf: "dialogf_en.tlk", dialogfAuto: false },
        ["One", "Two"],
        ["OneF", "TwoF"],
      ),
      makeParsedBundle(
        { id: 2, locale: "PL", dialog: "dialog_pl.tlk", dialogf: "", dialogfAuto: false },
        ["Jeden", "Dwa"],
      ),
    ];
    const localeColumns: LocaleColumn[] = [
      { field: "loc_en", title: "EN", locale: "EN", variant: "dialog" },
      { field: "loc_en_f", title: "EN F", locale: "EN", variant: "dialogf" },
      { field: "loc_pl", title: "PL", locale: "PL", variant: "dialog" },
    ];

    const rows = buildRowsFromParsedTlkBundles(bundles, localeColumns);
    expect(rows).toHaveLength(2);
    expect(rows[0].sourceEn).toBe("One");
    expect(rows[0].loc_en_f).toBe("OneF");
    expect(rows[1].loc_pl).toBe("Dwa");
  });

  it("builds empty rows when parsed bundle list is empty", () => {
    const rows = buildRowsFromParsedTlkBundles([], []);
    expect(rows).toEqual([]);
  });

  it("falls back sourceEn to first available locale when EN column is missing", () => {
    const bundles: ParsedTlkBundle[] = [
      makeParsedBundle(
        { id: 2, locale: "PL", dialog: "dialog_pl.tlk", dialogf: "", dialogfAuto: false },
        ["Jeden", "Dwa"],
      ),
    ];
    const localeColumns: LocaleColumn[] = [{ field: "loc_pl", title: "PL", locale: "PL", variant: "dialog" }];
    const rows = buildRowsFromParsedTlkBundles(bundles, localeColumns);
    expect(rows[0].sourceEn).toBe("Jeden");
  });
});

describe("XLSX mapping and runtime", () => {
  it("maps XLSX rows to TLK rows and exports back", () => {
    const sheetRows: XlsxRow[] = [
      { StrRef: 0, Source_EN: "Hello", EN: "Hello", PL: "Czesc", Status: "Draft" },
      { StrRef: 1, Source_EN: "World", EN: "World", PL: "Swiat", Status: "Draft" },
    ];
    const tlkRows = buildTlkRowsFromXlsxRows(sheetRows);

    expect(tlkRows).toHaveLength(2);
    expect(tlkRows[0].loc_pl).toBe("Czesc");

    const localeColumns: LocaleColumn[] = [
      { field: "loc_en", title: "EN", locale: "EN", variant: "dialog" },
      { field: "loc_pl", title: "PL", locale: "PL", variant: "dialog" },
    ];
    const exported = buildXlsxRowsForExport(tlkRows, localeColumns);
    expect(exported[0].PL).toBe("Czesc");
  });

  it("builds stable XLSX headers for export", () => {
    const localeColumns: LocaleColumn[] = [
      { field: "loc_en", title: "EN", locale: "EN", variant: "dialog" },
      { field: "loc_pl_f", title: "PL F", locale: "PL", variant: "dialogf" },
      { field: "loc_de_xlsx", title: "DE XLS", locale: "DE", variant: "xlsx-extra" },
    ];
    expect(buildXlsxHeaderForExport(localeColumns)).toEqual([
      "StrRef",
      "Source_EN",
      "Status",
      "EN",
      "PL_F",
      "DE_XLS",
    ]);
  });

  it("builds AOA export payload aligned with object export mapping", () => {
    const rows: TlkGridRow[] = [
      {
        id: 1,
        strRef: 1001,
        sourceEn: "Hello",
        context: "",
        status: "Draft",
        loc_en: "Hello",
        loc_pl_f: "Witaj",
        loc_de_xlsx: "Hallo",
      },
    ];
    const localeColumns: LocaleColumn[] = [
      { field: "loc_en", title: "EN", locale: "EN", variant: "dialog" },
      { field: "loc_pl_f", title: "PL F", locale: "PL", variant: "dialogf" },
      { field: "loc_de_xlsx", title: "DE XLS", locale: "DE", variant: "xlsx-extra" },
    ];

    const aoa = buildXlsxAoaForExport(rows, localeColumns);
    expect(aoa[0]).toEqual(["StrRef", "Source_EN", "Status", "EN", "PL_F", "DE_XLS"]);
    expect(aoa[1]).toEqual([1001, "Hello", "Draft", "Hello", "Witaj", "Hallo"]);

    const objectRows = buildXlsxRowsForExport(rows, localeColumns);
    expect(objectRows[0]).toEqual({
      StrRef: 1001,
      Source_EN: "Hello",
      Status: "Draft",
      EN: "Hello",
      PL_F: "Witaj",
      DE_XLS: "Hallo",
    });
  });

  it("maps special XLSX suffix columns", () => {
    const rows: XlsxRow[] = [{ StrRef: 7, Source: "A", FR_F: "AF", ES_XLS: "AE" }];
    const mapped = buildTlkRowsFromXlsxRows(rows);
    expect(mapped[0].loc_fr_f).toBe("AF");
    expect(mapped[0].loc_es_xlsx).toBe("AE");
  });

  it("filters bad StrRef rows and sorts output", () => {
    const rows: XlsxRow[] = [
      { StrRef: "x", Source_EN: "bad" },
      { StrRef: 2, Source_EN: "B" },
      { StrRef: 1, Source_EN: "A" },
      { Source_EN: "noRef" },
    ];
    const mapped = buildTlkRowsFromXlsxRows(rows);
    expect(mapped).toHaveLength(2);
    expect(mapped[0].strRef).toBe(1);
    expect(mapped[1].strRef).toBe(2);
  });

  it("parses locale columns from loc_* headers", () => {
    const parsed = localeColumnsFromXlsxColumns(["loc_pl", "loc_pl_f", "loc_de_xlsx", "Source_EN"]);
    expect(parsed.map((col) => col.variant)).toEqual(["dialog", "dialogf", "xlsx-extra"]);
  });

  it("parses locale columns from mapped rows", () => {
    const rows: TlkGridRow[] = [{ id: 1, strRef: 1, sourceEn: "A", context: "", status: "Draft", loc_en: "A", loc_pl_f: "B" }];
    const parsed = getParsedLocaleColumnsFromRows(rows);
    expect(parsed).toHaveLength(2);
    expect(parsed[1].variant).toBe("dialogf");
  });

  it("returns empty locale list when no rows", () => {
    expect(getParsedLocaleColumnsFromRows([])).toEqual([]);
  });

  it("sanitizes and names XLSX helpers", () => {
    expect(sanitizeXlsxCell("A\r\nB")).toBe("A\nB");
    expect(makeXlsxFileName()).toMatch(/^localization_sheet_v4_.*\.xlsx$/);
    expect(makeCsvFileName()).toBe("latest-localization.csv");
    expect(makeCsvFileName("dialog_en.tlk")).toBe("dialog_en.csv");
    expect(makeCsvFileName("C:\\packs\\dialog_pl.tlk")).toBe("dialog_pl.csv");
    expect(makeCsvFileName("custom sheet.xlsx")).toBe("custom sheet.csv");
    expect(makeCsvFileName("bad:name?.tlk")).toBe("bad_name_.csv");
  });

  it("detects publishable CSV changes based on source/locale values", () => {
    const localeColumns: LocaleColumn[] = [
      { field: "loc_en", title: "EN", locale: "EN", variant: "dialog" },
      { field: "loc_pl", title: "PL", locale: "PL", variant: "dialog" },
    ];
    const baseline = makeGridRows();
    const current = baseline.map((row) => ({ ...row }));
    current[1].loc_pl = "Nowy tekst";

    expect(
      hasCsvExportChanges({
        baselineRows: baseline,
        currentRows: current,
        localeColumns,
      }),
    ).toBe(true);
  });

  it("ignores status-only updates for publishable CSV changes", () => {
    const localeColumns: LocaleColumn[] = [{ field: "loc_en", title: "EN", locale: "EN", variant: "dialog" }];
    const baseline = makeGridRows();
    const current = baseline.map((row) => ({ ...row, status: "Validated" }));

    expect(
      hasCsvExportChanges({
        baselineRows: baseline,
        currentRows: current,
        localeColumns,
      }),
    ).toBe(false);
  });

  it("computes change summary for CSV export", () => {
    const localeColumns: LocaleColumn[] = [
      { field: "loc_en", title: "EN", locale: "EN", variant: "dialog" },
      { field: "loc_pl", title: "PL", locale: "PL", variant: "dialog" },
    ];
    const baseline: TlkGridRow[] = [
      { id: 0, strRef: 0, sourceEn: "Hello", context: "", status: "Draft", loc_en: "Hello", loc_pl: "Czesc" },
      { id: 1, strRef: 1, sourceEn: "World", context: "", status: "Draft", loc_en: "World", loc_pl: "Swiat" },
    ];
    const current: TlkGridRow[] = [
      { id: 0, strRef: 0, sourceEn: "Hello!", context: "", status: "Draft", loc_en: "Hello", loc_pl: "Czesc" },
      { id: 2, strRef: 2, sourceEn: "New", context: "", status: "Draft", loc_en: "New", loc_pl: "" },
    ];

    const summary = computeCsvExportChangeSummary({
      baselineRows: baseline,
      currentRows: current,
      localeColumns,
    });

    expect(summary.changedRows).toBe(3);
    expect(summary.addedRows).toBe(1);
    expect(summary.removedRows).toBe(1);
    expect(summary.sourceChangedRows).toBe(3);
    expect(summary.changedCells).toBeGreaterThanOrEqual(3);
    expect(summary.touchedLocales).toContain("EN");
    expect(summary.changedRowStrRefs).toEqual([0, 1, 2]);
  });

  it("treats row count mismatch as publishable change", () => {
    const localeColumns: LocaleColumn[] = [{ field: "loc_en", title: "EN", locale: "EN", variant: "dialog" }];
    const baseline = makeGridRows();
    const current = baseline.slice(0, 1);

    expect(
      hasCsvExportChanges({
        baselineRows: baseline,
        currentRows: current,
        localeColumns,
      }),
    ).toBe(true);
  });

  it("treats newly added locale column as publishable change even when cells are empty", () => {
    const baseline: TlkGridRow[] = [
      { id: 0, strRef: 0, sourceEn: "Hello", context: "", status: "Draft", loc_en: "Hello" },
    ];
    const current: TlkGridRow[] = [
      { id: 0, strRef: 0, sourceEn: "Hello", context: "", status: "Draft", loc_en: "Hello", loc_pl_xlsx: "" },
    ];
    const localeColumns: LocaleColumn[] = [
      { field: "loc_en", title: "EN", locale: "EN", variant: "dialog" },
      { field: "loc_pl_xlsx", title: "PL XLS", locale: "PL", variant: "xlsx-extra" },
    ];

    expect(
      hasCsvExportChanges({
        baselineRows: baseline,
        currentRows: current,
        localeColumns,
      }),
    ).toBe(true);
  });

  it("detects runtime support", () => {
    expect(hasXlsxRuntime(undefined)).toBe(false);
    expect(
      hasXlsxRuntime({
        read: () => ({ SheetNames: [], Sheets: {} }),
        utils: { sheet_to_json: () => [] },
      }),
    ).toBe(true);
  });

  it("parses workbook file through runtime adapter", async () => {
    const runtime = {
      read: () => ({ SheetNames: ["Sheet1"], Sheets: { Sheet1: {} } }),
      utils: { sheet_to_json: () => [{ StrRef: 1, Source_EN: "A" }] },
    };
    const file = {
      name: "sample.xlsx",
      arrayBuffer: async () => new ArrayBuffer(8),
    };

    const parsed = await parseWorkbookFile(file, runtime);
    expect(parsed.sheetName).toBe("Sheet1");
    expect(parsed.rows).toHaveLength(1);
  });

  it("throws for invalid workbook runtime states", async () => {
    const file = {
      name: "sample.xlsx",
      arrayBuffer: async () => new ArrayBuffer(8),
    };
    const runtimeNoSheets = {
      read: () => ({ SheetNames: [], Sheets: {} }),
      utils: { sheet_to_json: () => [{ StrRef: 1, Source_EN: "A" }] },
    };
    const runtimeNoRows = {
      read: () => ({ SheetNames: ["S"], Sheets: { S: {} } }),
      utils: { sheet_to_json: () => [] },
    };

    await expect(parseWorkbookFile(undefined as never, runtimeNoSheets)).rejects.toThrow("Select XLSX file first");
    await expect(parseWorkbookFile(file, runtimeNoSheets)).rejects.toThrow("no sheets");
    await expect(parseWorkbookFile(file, runtimeNoRows)).rejects.toThrow("no data rows");
    await expect(parseWorkbookFile(file, {} as never)).rejects.toThrow("runtime failed");
  });

  it("validates xlsx schema happy path", () => {
    const result = validateXlsxSchema([{ StrRef: 1, Source_EN: "X", loc_en: "X" }], "sheet");
    expect(result.ok).toBe(true);
  });

  it("validates xlsx schema error/warning cases", () => {
    const empty = validateXlsxSchema([], "sheet");
    expect(empty.ok).toBe(false);

    const bad = validateXlsxSchema(
      [{ StrRef: "NaN", Source: "", loc_en: "A", loc_EN: "B" }, { StrRef: "x", Source: "A" }],
      "sheet",
    );
    expect(bad.ok).toBe(false);
    expect(bad.issues.some((i) => i.message.includes("invalid StrRef"))).toBe(true);
    expect(bad.issues.some((i) => i.message.includes("duplicate locale columns"))).toBe(true);

    const noLocale = validateXlsxSchema([{ StrRef: 1, Source_EN: "A" }], "sheet");
    expect(noLocale.issues.some((i) => i.severity === "warning")).toBe(true);
  });
});

describe("GitHub publish", () => {
  it("rejects when token is missing", async () => {
    await expect(
      publishCsvToGitHub({
        token: "",
        repoFullName: "enonwow/nwn-localization-test",
        baseBranch: "main",
        csvFolder: "csv-latest",
        fileName: "dialog_en.csv",
        csvBytes: new TextEncoder().encode("A,B\r\n1,2"),
      }),
    ).rejects.toThrow("token");
  });

  it("rejects invalid repo format", async () => {
    await expect(
      publishCsvToGitHub({
        token: "test-token",
        repoFullName: "invalid-repo-name",
        baseBranch: "main",
        csvFolder: "csv-latest",
        fileName: "dialog_en.csv",
        csvBytes: new TextEncoder().encode("A,B\r\n1,2"),
      }),
    ).rejects.toThrow("owner/repo");
  });

  it("creates branch, commits CSV and opens PR via GitHub API", async () => {
    const originalFetch = globalThis.fetch;
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      const method = String(init?.method || "GET").toUpperCase();

      if (url.endsWith("/git/ref/heads/main") && method === "GET") {
        return new Response(JSON.stringify({ object: { sha: "base-sha-1" } }), { status: 200 });
      }
      if (url.endsWith("/git/refs") && method === "POST") {
        return new Response(JSON.stringify({ ref: "refs/heads/tlk-forge/test" }), { status: 201 });
      }
      if (url.includes("/contents/csv-latest/dialog_en.csv") && method === "GET") {
        return new Response(JSON.stringify({ message: "Not Found" }), { status: 404 });
      }
      if (url.includes("/contents/csv-latest/dialog_en.csv") && method === "PUT") {
        return new Response(JSON.stringify({ commit: { sha: "commit-sha-1" } }), { status: 200 });
      }
      if (url.endsWith("/pulls") && method === "POST") {
        return new Response(
          JSON.stringify({ html_url: "https://github.com/enonwow/nwn-localization-test/pull/1" }),
          { status: 201 },
        );
      }
      return new Response(JSON.stringify({ message: `Unhandled ${method} ${url}` }), { status: 500 });
    });

    (globalThis as { fetch: typeof fetch }).fetch = fetchMock as unknown as typeof fetch;
    try {
      const result = await publishCsvToGitHub({
        token: "test-token",
        repoFullName: "enonwow/nwn-localization-test",
        baseBranch: "main",
        csvFolder: "csv-latest",
        fileName: "dialog_en.csv",
        csvBytes: new TextEncoder().encode("A,B\r\n1,2"),
      });

      expect(result.prUrl).toBe("https://github.com/enonwow/nwn-localization-test/pull/1");
      expect(result.commitSha).toBe("commit-sha-1");
      expect(result.filePath).toBe("csv-latest/dialog_en.csv");
      expect(fetchMock).toHaveBeenCalledTimes(5);
    } finally {
      (globalThis as { fetch: typeof fetch }).fetch = originalFetch;
    }
  });

  it("downloads CSV from repository branch", async () => {
    const originalFetch = globalThis.fetch;
    const csvText = "StrRef,Source_EN,Status\r\n1,Hello,Draft";
    const fetchMock = vi.fn(async (input: RequestInfo | URL) => {
      const url = String(input);
      if (url.includes("/contents/csv-latest/dialog_en.csv?ref=main")) {
        return new Response(new TextEncoder().encode(csvText), { status: 200 });
      }
      return new Response("not-found", { status: 404 });
    });

    (globalThis as { fetch: typeof fetch }).fetch = fetchMock as unknown as typeof fetch;
    try {
      const result = await fetchRepoFileFromGitHub({
        repoFullName: "enonwow/nwn-localization-test",
        branch: "main",
        filePath: "csv-latest/dialog_en.csv",
      });

      expect(result.fileName).toBe("dialog_en.csv");
      expect(result.filePath).toBe("csv-latest/dialog_en.csv");
      expect(new TextDecoder().decode(new Uint8Array(result.bytes))).toBe(csvText);
    } finally {
      (globalThis as { fetch: typeof fetch }).fetch = originalFetch;
    }
  });
});

describe("Validation orchestration", () => {
  it("validates TLK bundles and catches structural errors", () => {
    const empty = validateTlkBundles([]);
    expect(empty.ok).toBe(false);

    const invalid = validateTlkBundles([
      { id: 1, locale: "", dialog: "dialog.txt", dialogf: "dialogf.txt", dialogfAuto: false },
      { id: 2, locale: "PL", dialog: "", dialogf: "", dialogfAuto: false },
    ]);
    expect(invalid.ok).toBe(false);
    expect(invalid.issues.length).toBeGreaterThan(0);
  });

  it("applies fallback only when dialogf is missing", () => {
    const result = applyDialogfFallbacks([
      { id: 1, locale: "EN", dialog: "dialog_en.tlk", dialogf: "", dialogfAuto: false },
      { id: 2, locale: "PL", dialog: "dialog_pl.tlk", dialogf: "dialogf_pl.tlk", dialogfAuto: false },
    ]);
    expect(result.fallbackCount).toBe(1);
    expect(result.bundles[0].dialogfAuto).toBe(true);
    expect(result.bundles[1].dialogfAuto).toBe(false);
  });

  it("validates schema from workbook rows and combines results", () => {
    const workbook = { workbook: {}, sheetName: "Main", rows: [] };
    const result = validateSchemaFromXlsxWorkbook(workbook, [{ strRef: 1, sourceEn: "A" }]);
    expect(result.issues.some((i) => i.message.includes("missing required column StrRef"))).toBe(true);
    expect(result.issues.some((i) => i.message.includes("No locale columns"))).toBe(true);

    const first: ValidationResult = { ok: true, issues: [{ severity: "warning", message: "w" }] };
    const second: ValidationResult = { ok: true, issues: [{ severity: "error", message: "e" }] };
    const combined = combineValidationResult(first, second);
    expect(combined.ok).toBe(false);
    expect(isValidationWarning("warning")).toBe(true);
    expect(isValidationWarning("error")).toBe(false);
  });
});

describe("Diff behavior", () => {
  it("computes added, removed and conflicts", () => {
    const localeColumns: LocaleColumn[] = [{ field: "loc_pl", title: "PL", locale: "PL", variant: "dialog" }];
    const baseline: TlkGridRow[] = [
      { id: 0, strRef: 0, sourceEn: "A", context: "", status: "Draft", loc_pl: "A" },
      { id: 1, strRef: 1, sourceEn: "B", context: "", status: "Draft", loc_pl: "B {TOKEN}" },
      { id: 2, strRef: 2, sourceEn: "C", context: "", status: "Draft", loc_pl: "C" },
    ];
    const current: TlkGridRow[] = [
      { id: 0, strRef: 0, sourceEn: "A", context: "", status: "Draft", loc_pl: "A+" },
      { id: 1, strRef: 1, sourceEn: "B", context: "", status: "Draft", loc_pl: "B {OTHER}" },
      { id: 3, strRef: 3, sourceEn: "D", context: "", status: "Draft", loc_pl: "D" },
    ];

    const result = computeDiff({ baselineRows: baseline, currentRows: current, localeColumns, previewLimit: 20 });
    expect(result.counts.changed).toBe(4);
    expect(result.counts.added).toBe(1);
    expect(result.counts.removed).toBe(1);
    expect(result.counts.conflicts).toBe(1);
  });

  it("renders markdown report", () => {
    const localeColumns: LocaleColumn[] = [{ field: "loc_pl", title: "PL", locale: "PL", variant: "dialog" }];
    const baseline: TlkGridRow[] = [{ id: 0, strRef: 0, sourceEn: "A", context: "", status: "Draft", loc_pl: "A" }];
    const current: TlkGridRow[] = [{ id: 0, strRef: 0, sourceEn: "A", context: "", status: "Draft", loc_pl: "B" }];
    const result = computeDiff({ baselineRows: baseline, currentRows: current, localeColumns });
    const md = renderDiffMarkdown(result.report);
    expect(md).toContain("# TLK Diff Report");
    expect(md).toContain("| StrRef | Locale | Before | After | Status |");
  });
});
