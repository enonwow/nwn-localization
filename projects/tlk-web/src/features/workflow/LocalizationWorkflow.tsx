import { useCallback, useMemo, useRef, useState } from "react";
import type { GridApi } from "ag-grid-community";
import * as XLSX from "xlsx";

import ScopeSwitcher from "../../components/ScopeSwitcher";
import StepStrip from "../../components/StepStrip";
import LocalizationGrid from "../../components/LocalizationGrid";
import { computeDiff, renderDiffMarkdown } from "../../lib/diff";
import {
  buildRowsFromParsedTlkBundles,
  buildSingleTlkBinaryFromColumn,
  makeTlkFileName,
  parseSingleTlkBuffer,
  quickChecksumHex,
  TLK_LOCALE_TO_LANGUAGE_ID,
} from "../../lib/tlk";
import { localeCodeToFieldToken, normalizeLocaleCode, type LocaleColumn, type ParsedTlkBundle, type TlkBundleConfig, type TlkGridRow, type TlkDiffReport } from "../../lib/types";
import { applyDialogfFallbacks, validateTlkBundles } from "../../lib/validation";
import {
  buildTlkRowsFromXlsxRows,
  getParsedLocaleColumnsFromRows,
  makeCsvFileName,
  parseWorkbookFile,
} from "../../lib/xlsx";
import { computeStepUiStates } from "../../lib/workflowProgress";
import type { ImportMode, WorkflowScope, WorkflowStep } from "./types";

type ArtifactRow = {
  file: string;
  checksum: string;
};

type ValidationFeedback = {
  level: "ok" | "warning" | "error";
  message: string;
  canProceed: boolean;
  warningCount: number;
  errorCount: number;
  warningSamples: Array<{ strRef: number; reason: string }>;
  errorSamples: Array<{ strRef: number; reason: string }>;
};

type BundleFileSlot = {
  dialog: File | null;
  dialogf: File | null;
};

type BundleFileKind = keyof BundleFileSlot;

type CsvExportWorkerResponse =
  | {
      type: "progress";
      doneRows: number;
      totalRows: number;
      progress: number;
    }
  | {
      type: "result";
      fileName: string;
      bytes: ArrayBuffer;
    }
  | {
      type: "error";
      error: string;
    };

const TLK_LOCALE_OPTIONS = ["EN", "PL", "DE", "FR", "ES", "IT", "PT-BR", "RU"];

const STEP_MAP: Record<WorkflowScope, WorkflowStep[]> = {
  exchange: [
    { title: "Import TLK", sub: "Parse + schema map" },
    { title: "Edit TLK", sub: "Manual entry fixes" },
    { title: "Publish", sub: "Generate CSV + Open PR" },
  ],
  rebuild: [
    { title: "Import CSV", sub: "Load approved file" },
    { title: "Rebuild TLK", sub: "Generate TLK/TLKF" },
  ],
};

const DEFAULT_BUNDLES: TlkBundleConfig[] = [
  { id: 1, locale: "EN", dialog: "dialog_en.tlk", dialogf: "dialogf_en.tlk", dialogfAuto: false },
];

function isTlkFile(file: File): boolean {
  return /\.tlk$/i.test(String(file.name || "").trim());
}

function cloneRows(rows: readonly TlkGridRow[]): TlkGridRow[] {
  return rows.map((row) => ({ ...row }));
}

function buildLoadedLocaleColumns(bundles: readonly TlkBundleConfig[]): LocaleColumn[] {
  const used = new Set<string>();
  const result: LocaleColumn[] = [];

  bundles.forEach((bundle, index) => {
    const locale = String(bundle.locale || "").toUpperCase().trim();
    if (!locale) return;
    const token = locale.toLowerCase().replace(/[^a-z0-9]+/g, "_") || `locale_${index + 1}`;
    const dialogField = `loc_${token}`;
    if (!used.has(dialogField)) {
      used.add(dialogField);
      result.push({
        field: dialogField,
        title: locale,
        locale,
        variant: "dialog",
      });
    }

    const hasDialogf = String(bundle.dialogf || "").trim().length > 0 && !bundle.dialogfAuto;
    if (hasDialogf) {
      const dialogfField = `loc_${token}_f`;
      if (!used.has(dialogfField)) {
        used.add(dialogfField);
        result.push({
          field: dialogfField,
          title: `${locale} F`,
          locale,
          variant: "dialogf",
        });
      }
    }
  });

  return result;
}

function normalizeExtraLocales(loaded: readonly LocaleColumn[], extraLocales: readonly string[]): string[] {
  const loadedSet = new Set(loaded.map((col) => col.locale));
  const seen = new Set<string>();

  return extraLocales
    .map((locale) => String(locale || "").toUpperCase().trim())
    .filter((locale) => {
      if (!locale || loadedSet.has(locale) || seen.has(locale)) return false;
      seen.add(locale);
      return true;
    });
}

function buildEffectiveLocaleColumns(bundles: readonly TlkBundleConfig[], extraLocales: readonly string[]): LocaleColumn[] {
  const loadedColumns = buildLoadedLocaleColumns(bundles);
  const normalizedExtras = normalizeExtraLocales(loadedColumns, extraLocales);
  const extraColumns: LocaleColumn[] = normalizedExtras.map((locale) => {
    const token = locale.toLowerCase().replace(/[^a-z0-9]+/g, "_");
    return {
      field: `loc_${token}_xlsx`,
      title: `${locale} XLS`,
      locale,
      variant: "xlsx-extra",
    };
  });
  return [...loadedColumns, ...extraColumns];
}

async function saveBytesToDisk(bytes: Uint8Array, fileName: string, mimeType: string): Promise<string> {
  const payload = new Uint8Array(bytes.byteLength);
  payload.set(bytes);
  const blob = new Blob([payload], { type: mimeType });
  if ("showSaveFilePicker" in window) {
    const picker = window as unknown as {
      showSaveFilePicker: (options: { suggestedName: string }) => Promise<FileSystemFileHandle>;
    };
    const handle = await picker.showSaveFilePicker({ suggestedName: fileName });
    const writable = await handle.createWritable();
    await writable.write(blob);
    await writable.close();
    return fileName;
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
  return fileName;
}

async function yieldToUiFrame(): Promise<void> {
  await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()));
}

function createCsvExportTask(
  rows: TlkGridRow[],
  localeColumns: LocaleColumn[],
  fileName: string,
  onProgress?: (payload: { doneRows: number; totalRows: number; progress: number }) => void,
) {
  const worker = new Worker(new URL("../../workers/csvExportWorker.ts", import.meta.url), { type: "module" });
  let isSettled = false;
  let rejectPromise: ((reason?: unknown) => void) | null = null;
  const promise = new Promise<{ fileName: string; bytes: Uint8Array }>((resolve, reject) => {
    rejectPromise = reject;
    worker.onmessage = (event: MessageEvent<CsvExportWorkerResponse>) => {
      const payload = event.data;
      if (isSettled) return;
      if (payload.type === "progress") {
        onProgress?.({ doneRows: payload.doneRows, totalRows: payload.totalRows, progress: payload.progress });
        return;
      }
      if (payload.type === "error") {
        isSettled = true;
        reject(new Error(payload.error || "CSV export worker failed."));
        return;
      }
      if (payload.type === "result") {
        isSettled = true;
        resolve({ fileName: payload.fileName, bytes: new Uint8Array(payload.bytes) });
      }
    };
    worker.onerror = (event) => {
      if (isSettled) return;
      isSettled = true;
      reject(new Error(event.message || "CSV export worker crashed."));
    };
  });

  worker.postMessage({ rows, localeColumns, fileName });
  const cancel = () => {
    if (isSettled) return;
    isSettled = true;
    worker.terminate();
    rejectPromise?.(new Error("Export canceled."));
  };
  return { worker, promise, cancel };
}

function statusBadgeClass(value: string): string {
  const normalized = value.toLowerCase();
  if (normalized === "validated") return "status-ok";
  if (normalized === "error") return "status-error";
  if (normalized === "needs qa" || normalized === "mt draft") return "status-warn";
  return "status-pending";
}

function bundleIdKey(id: number | string | undefined): string {
  return String(id ?? "");
}

const LocalizationWorkflow = () => {
  const [scope, setScope] = useState<WorkflowScope>("exchange");
  const [stepIndex, setStepIndex] = useState(0);
  const [importMode, setImportMode] = useState<ImportMode>("tlk");
  const [tlkBundles, setTlkBundles] = useState<TlkBundleConfig[]>(DEFAULT_BUNDLES);
  const [bundleSeq, setBundleSeq] = useState(DEFAULT_BUNDLES.length + 1);
  const [bundleFiles, setBundleFiles] = useState<Record<string, BundleFileSlot>>({});
  const [extraXlsxLocales, setExtraXlsxLocales] = useState<string[]>([]);
  const [sourceXlsxFile, setSourceXlsxFile] = useState<File | null>(null);
  const [mergedXlsxFile, setMergedXlsxFile] = useState<File | null>(null);
  const [rows, setRows] = useState<TlkGridRow[]>([]);
  const [baselineRows, setBaselineRows] = useState<TlkGridRow[]>([]);
  const [localeColumns, setLocaleColumns] = useState<LocaleColumn[]>([]);
  const [sourceLoaded, setSourceLoaded] = useState(false);
  const [validated, setValidated] = useState(false);
  const [exported, setExported] = useState(false);
  const [imported, setImported] = useState(false);
  const [built, setBuilt] = useState(false);
  const [diffReport, setDiffReport] = useState<TlkDiffReport | null>(null);
  const [artifacts, setArtifacts] = useState<ArtifactRow[]>([]);
  const [pageSize, setPageSize] = useState(25);
  const [, setStatusMessage] = useState("Ready. Aurora-first workflow is active.");
  const [validationFeedback, setValidationFeedback] = useState<ValidationFeedback | null>(null);
  const [isExportingXlsx, setIsExportingXlsx] = useState(false);
  const [exportProgress, setExportProgress] = useState(0);
  const [exportProgressLabel, setExportProgressLabel] = useState("");
  const [lastExport, setLastExport] = useState<{ fileName: string; bytes: Uint8Array } | null>(null);
  const [exportRunMode, setExportRunMode] = useState<"generate" | "publish">("generate");
  const [csvExportFileName, setCsvExportFileName] = useState("latest-localization.csv");
  const [activeDropKey, setActiveDropKey] = useState<string | null>(null);
  const [prUrl, setPrUrl] = useState("");
  const [gridApi, setGridApi] = useState<GridApi<TlkGridRow> | null>(null);
  const exportWorkerRef = useRef<Worker | null>(null);
  const exportCancelRef = useRef<(() => void) | null>(null);
  const bundleInputRefs = useRef<Record<string, HTMLInputElement | null>>({});

  const steps = STEP_MAP[scope];
  const loadedLocaleColumns = useMemo(() => buildLoadedLocaleColumns(tlkBundles), [tlkBundles]);
  const localeBaseForExtras = useMemo(
    () =>
      sourceLoaded && localeColumns.length > 0
        ? localeColumns.filter((col) => col.variant !== "xlsx-extra")
        : loadedLocaleColumns,
    [loadedLocaleColumns, localeColumns, sourceLoaded],
  );
  const normalizedExtraLocales = useMemo(
    () => normalizeExtraLocales(localeBaseForExtras, extraXlsxLocales),
    [extraXlsxLocales, localeBaseForExtras],
  );
  const availableExtraLocales = useMemo(() => {
    const loadedSet = new Set(localeBaseForExtras.map((col) => col.locale));
    const selectedSet = new Set(normalizedExtraLocales);
    return TLK_LOCALE_OPTIONS.filter((locale) => !loadedSet.has(locale) && !selectedSet.has(locale));
  }, [localeBaseForExtras, normalizedExtraLocales]);

  const totalLocaleCells = useMemo(
    () => rows.length * Math.max(1, localeColumns.length),
    [localeColumns.length, rows.length],
  );

  const filledLocaleCells = useMemo(() => {
    return rows.reduce((sum, row) => {
      const count = localeColumns.filter((col) => String(row[col.field] || "").trim().length > 0).length;
      return sum + count;
    }, 0);
  }, [localeColumns, rows]);

  const coverage = totalLocaleCells > 0 ? (filledLocaleCells / totalLocaleCells) * 100 : 0;
  const qaWarnings = useMemo(
    () => rows.filter((row) => String(row.status || "").toLowerCase() === "needs qa").length,
    [rows],
  );
  const blockingErrors = useMemo(
    () => rows.filter((row) => String(row.status || "").toLowerCase() === "error").length,
    [rows],
  );

  const resetFlowState = useCallback(() => {
    setRows([]);
    setBaselineRows([]);
    setLocaleColumns([]);
    setSourceLoaded(false);
    setValidated(false);
    setExported(false);
    setImported(false);
    setBuilt(false);
    setDiffReport(null);
    setArtifacts([]);
    setBundleFiles({});
    setValidationFeedback(null);
    if (exportWorkerRef.current) {
      exportWorkerRef.current.terminate();
      exportWorkerRef.current = null;
    }
    exportCancelRef.current = null;
    setIsExportingXlsx(false);
    setExportProgress(0);
    setExportProgressLabel("");
    setLastExport(null);
    setCsvExportFileName("latest-localization.csv");
    setActiveDropKey(null);
  }, []);

  const onScopeChange = useCallback(
    (nextScope: WorkflowScope) => {
      setScope(nextScope);
      setStepIndex(0);
      setImportMode(nextScope === "exchange" ? "tlk" : "xlsx");
      resetFlowState();
      setStatusMessage(
        nextScope === "exchange"
          ? "Exchange scope selected. Start with source import."
          : "Rebuild scope selected. Import approved CSV.",
      );
    },
    [resetFlowState],
  );

  const onGridRowsChange = useCallback((nextRows: TlkGridRow[]) => {
    setRows(nextRows);
    setValidated(false);
    setExported(false);
    setBuilt(false);
    setLastExport(null);
    setValidationFeedback(null);
  }, []);

  const onGridApiReady = useCallback((api: GridApi<TlkGridRow>) => {
    setGridApi(api);
  }, []);

  const applyQaValidation = useCallback(() => {
    if (!sourceLoaded || rows.length === 0) {
      setValidationFeedback({
        level: "error",
        message: "Load source first, then validate.",
        canProceed: false,
        warningCount: 0,
        errorCount: 0,
        warningSamples: [],
        errorSamples: [],
      });
      setStatusMessage("Load source first.");
      return;
    }

    let tokenMismatch = 0;
    let emptyRows = 0;
    const warningSamples: Array<{ strRef: number; reason: string }> = [];
    const errorSamples: Array<{ strRef: number; reason: string }> = [];
    const SAMPLE_LIMIT = 20;
    const nextRows = rows.map((row) => {
      const missingColumns = localeColumns.filter((col) => !String(row[col.field] || "").trim());
      const hasEmpty = missingColumns.length > 0;
      if (hasEmpty) {
        emptyRows += 1;
        if (warningSamples.length < SAMPLE_LIMIT) {
          warningSamples.push({
            strRef: row.strRef,
            reason: `Missing translation in: ${missingColumns.map((col) => col.title).slice(0, 4).join(", ")}`,
          });
        }
        return { ...row, status: "Needs QA" };
      }
      const tokenIssueColumns = localeColumns.filter((col) => {
        const value = String(row[col.field] || "");
        return value.includes("{PLAYER}") && !value.includes("{PLAYER_NAME}");
      });
      const hasTokenIssue = tokenIssueColumns.length > 0;
      if (hasTokenIssue) {
        tokenMismatch += 1;
        if (errorSamples.length < SAMPLE_LIMIT) {
          errorSamples.push({
            strRef: row.strRef,
            reason: `Token mismatch in: ${tokenIssueColumns.map((col) => col.title).slice(0, 4).join(", ")}`,
          });
        }
        return { ...row, status: "Error" };
      }
      return { ...row, status: "Validated" };
    });

    setRows(nextRows);
    const hasBlockingErrors = tokenMismatch > 0;
    const isValidForNextStep = !hasBlockingErrors;
    setValidated(isValidForNextStep);
    const validatedCount = nextRows.length - emptyRows - tokenMismatch;
    setValidationFeedback(
      isValidForNextStep && emptyRows === 0
        ? {
            level: "ok",
            message: `Validation passed: ${validatedCount}/${nextRows.length} rows validated.`,
            canProceed: true,
            warningCount: emptyRows,
            errorCount: tokenMismatch,
            warningSamples,
            errorSamples,
          }
        : {
            level: hasBlockingErrors ? "error" : "warning",
            message: hasBlockingErrors
              ? `Validation blocked: Needs QA=${emptyRows}, Errors=${tokenMismatch}, Validated=${validatedCount}.`
              : `Validation warnings: Needs QA=${emptyRows}, Errors=${tokenMismatch}, Validated=${validatedCount}.`,
            canProceed: !hasBlockingErrors,
            warningCount: emptyRows,
            errorCount: tokenMismatch,
            warningSamples,
            errorSamples,
          },
    );
    setStatusMessage(
      hasBlockingErrors
        ? `Validation issues: empty rows=${emptyRows}, token conflicts=${tokenMismatch}.`
        : emptyRows > 0
          ? `Validation warnings: empty rows=${emptyRows}.`
          : `Validation passed. ${nextRows.length}/${nextRows.length} rows validated.`,
    );
  }, [localeColumns, rows, sourceLoaded]);

  const addBundle = useCallback(() => {
    const used = new Set(tlkBundles.map((bundle) => bundle.locale));
    const nextLocale = TLK_LOCALE_OPTIONS.find((locale) => !used.has(locale)) || "EN";
    const suffix = nextLocale.toLowerCase().replace(/[^a-z0-9]/g, "");
    const nextBundle: TlkBundleConfig = {
      id: bundleSeq,
      locale: nextLocale,
      dialog: `dialog_${suffix}.tlk`,
      dialogf: `dialogf_${suffix}.tlk`,
      dialogfAuto: false,
    };
    setBundleSeq((prev) => prev + 1);
    setTlkBundles((prev) => [...prev, nextBundle]);
    setSourceLoaded(false);
  }, [bundleSeq, tlkBundles]);

  const updateBundle = useCallback((id: number | string | undefined, patch: Partial<TlkBundleConfig>) => {
    setTlkBundles((prev) =>
      prev.map((bundle) => {
        if (bundle.id !== id) return bundle;
        return { ...bundle, ...patch };
      }),
    );
    setSourceLoaded(false);
  }, []);

  const removeBundle = useCallback((id: number | string | undefined) => {
    setTlkBundles((prev) => prev.filter((bundle) => bundle.id !== id));
    setBundleFiles((prev) => {
      const next = { ...prev };
      delete next[bundleIdKey(id)];
      return next;
    });
    setSourceLoaded(false);
  }, []);

  const updateBundleFile = useCallback(
    (id: number | string | undefined, kind: keyof BundleFileSlot, file: File | null) => {
      const key = bundleIdKey(id);
      setBundleFiles((prev) => {
        const current = prev[key] ?? { dialog: null, dialogf: null };
        return {
          ...prev,
          [key]: { ...current, [kind]: file },
        };
      });

      if (file) {
        if (kind === "dialog") {
          updateBundle(id, { dialog: file.name });
        } else {
          updateBundle(id, { dialogf: file.name, dialogfAuto: false });
        }
      } else {
        setSourceLoaded(false);
      }
    },
    [updateBundle],
  );

  const getBundleDropKey = useCallback((id: number | string | undefined, kind: BundleFileKind) => {
    return `${bundleIdKey(id)}:${kind}`;
  }, []);

  const setBundleInputRef = useCallback(
    (id: number | string | undefined, kind: BundleFileKind, input: HTMLInputElement | null) => {
      bundleInputRefs.current[getBundleDropKey(id, kind)] = input;
    },
    [getBundleDropKey],
  );

  const getBundleFile = useCallback(
    (id: number | string | undefined, kind: BundleFileKind): File | null => {
      const stateFile = (bundleFiles[bundleIdKey(id)] ?? { dialog: null, dialogf: null })[kind];
      if (stateFile) return stateFile;
      const input = bundleInputRefs.current[getBundleDropKey(id, kind)];
      return input?.files?.[0] ?? null;
    },
    [bundleFiles, getBundleDropKey],
  );

  const onBundleDragOver = useCallback(
    (event: React.DragEvent<HTMLElement>, id: number | string | undefined, kind: BundleFileKind) => {
      event.preventDefault();
      event.stopPropagation();
      const key = getBundleDropKey(id, kind);
      if (activeDropKey !== key) {
        setActiveDropKey(key);
      }
      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = "copy";
      }
    },
    [activeDropKey, getBundleDropKey],
  );

  const onBundleDragLeave = useCallback(
    (event: React.DragEvent<HTMLElement>, id: number | string | undefined, kind: BundleFileKind) => {
      event.preventDefault();
      event.stopPropagation();
      const key = getBundleDropKey(id, kind);
      setActiveDropKey((prev) => (prev === key ? null : prev));
    },
    [getBundleDropKey],
  );

  const onBundleDrop = useCallback(
    (event: React.DragEvent<HTMLElement>, id: number | string | undefined, kind: BundleFileKind) => {
      event.preventDefault();
      event.stopPropagation();

      const key = getBundleDropKey(id, kind);
      setActiveDropKey((prev) => (prev === key ? null : prev));

      const files = event.dataTransfer?.files;
      if (!files || files.length === 0) {
        setStatusMessage("No file dropped.");
        return;
      }

      const file = files[0];
      if (!isTlkFile(file)) {
        setStatusMessage("Drop a .tlk file.");
        return;
      }

      updateBundleFile(id, kind, file);
      if (files.length > 1) {
        setStatusMessage(`Multiple files dropped, using first: ${file.name}`);
        return;
      }
      setStatusMessage(`Loaded from drop: ${file.name}`);
    },
    [getBundleDropKey, updateBundleFile],
  );

  const syncExtraLocalesInGrid = useCallback(
    (nextExtraLocales: string[]) => {
      if (!sourceLoaded || localeColumns.length === 0) return;

      const baseColumns = localeColumns.filter((col) => col.variant !== "xlsx-extra");
      const normalized = normalizeExtraLocales(baseColumns, nextExtraLocales);
      const extraColumns: LocaleColumn[] = normalized.map((locale) => ({
        field: `loc_${localeCodeToFieldToken(locale)}_xlsx`,
        title: `${locale} XLS`,
        locale,
        variant: "xlsx-extra",
      }));
      const nextColumns = [...baseColumns, ...extraColumns];

      setLocaleColumns(nextColumns);
      setRows((prev) =>
        prev.map((row) => {
          const nextRow = { ...row };
          nextColumns.forEach((col) => {
            if (nextRow[col.field] === undefined) {
              nextRow[col.field] = "";
            }
          });
          return nextRow;
        }),
      );
      setValidated(false);
      setExported(false);
      setBuilt(false);
      setLastExport(null);
      setValidationFeedback(null);
    },
    [localeColumns, sourceLoaded],
  );

  const addExtraLocale = useCallback((locale: string) => {
    const normalized = normalizeLocaleCode(locale);
    if (!normalized) return;
    setExtraXlsxLocales((prev) => {
      if (prev.includes(normalized)) return prev;
      const next = [...prev, normalized];
      syncExtraLocalesInGrid(next);
      return next;
    });
  }, [syncExtraLocalesInGrid]);

  const removeExtraLocale = useCallback((locale: string) => {
    const normalized = normalizeLocaleCode(locale);
    setExtraXlsxLocales((prev) => {
      const next = prev.filter((item) => item !== normalized);
      syncExtraLocalesInGrid(next);
      return next;
    });
  }, [syncExtraLocalesInGrid]);

  const loadFromTlk = useCallback(async () => {
    const bundlesFromFiles = tlkBundles.map((bundle) => {
      const dialogFile = getBundleFile(bundle.id, "dialog");
      const dialogfFile = getBundleFile(bundle.id, "dialogf");
      return {
        ...bundle,
        dialog: dialogFile?.name || "",
        dialogf: dialogfFile?.name || "",
        dialogfAuto: false,
      };
    });

    const validation = validateTlkBundles(bundlesFromFiles);
    if (!validation.ok) {
      setStatusMessage(validation.issues[0]?.message || "TLK bundle validation failed.");
      return;
    }

    const fallback = applyDialogfFallbacks(validation.normalized);
    setTlkBundles(fallback.bundles);

    const parsedBundles: ParsedTlkBundle[] = [];
    for (let i = 0; i < fallback.bundles.length; i += 1) {
      const bundle = fallback.bundles[i];
      const dialogFile = getBundleFile(bundle.id, "dialog");
      if (!dialogFile) {
        setStatusMessage(`Locale ${bundle.locale}: select dialog.tlk file.`);
        return;
      }
      const parsedDialog = parseSingleTlkBuffer(await dialogFile.arrayBuffer(), dialogFile.name);

      let parsedDialogf = null;
      if (!bundle.dialogfAuto && String(bundle.dialogf || "").trim().length > 0) {
        const dialogfFile = getBundleFile(bundle.id, "dialogf");
        if (!dialogfFile) {
          setStatusMessage(`Locale ${bundle.locale}: select dialogf.tlk file or use fallback.`);
          return;
        }
        parsedDialogf = parseSingleTlkBuffer(await dialogfFile.arrayBuffer(), dialogfFile.name);
        if (parsedDialogf.entryCount !== parsedDialog.entryCount) {
          setStatusMessage(`Entry mismatch for locale ${bundle.locale}: dialog vs dialogf.`);
          return;
        }
      }
      parsedBundles.push({ bundle, parsed: parsedDialog, dialogfParsed: parsedDialogf });
    }

    const baseCount = parsedBundles[0]?.parsed.entryCount || 0;
    const mismatch = parsedBundles.find((item) => item.parsed.entryCount !== baseCount);
    if (mismatch) {
      setStatusMessage("Entry count mismatch across locale packs.");
      return;
    }

    const effectiveColumns = buildEffectiveLocaleColumns(fallback.bundles, normalizedExtraLocales);
    const builtRows = buildRowsFromParsedTlkBundles(parsedBundles, effectiveColumns);
    const primarySourceName = parsedBundles[0]?.parsed.fileName || fallback.bundles[0]?.dialog || "";
    setCsvExportFileName(makeCsvFileName(primarySourceName));
    setLocaleColumns(effectiveColumns);
    setRows(cloneRows(builtRows));
    setBaselineRows(cloneRows(builtRows));
    setSourceLoaded(true);
    setValidated(false);
    setExported(false);
    setBuilt(false);
    setLastExport(null);
    setStatusMessage(
      fallback.fallbackCount > 0
        ? `Loaded ${builtRows.length} rows with dialogf fallback for ${fallback.fallbackCount} pack(s).`
        : `Loaded ${builtRows.length} rows from TLK bundles.`,
    );
  }, [getBundleFile, normalizedExtraLocales, tlkBundles]);

  const loadFromXlsx = useCallback(
    async (file: File) => {
      const workbook = await parseWorkbookFile(file, XLSX);
      const builtRows = buildTlkRowsFromXlsxRows(workbook.rows);
      if (builtRows.length === 0) {
        setStatusMessage("CSV has no StrRef rows.");
        return false;
      }
      const parsedColumns = getParsedLocaleColumnsFromRows(builtRows);
      setCsvExportFileName(makeCsvFileName(file.name));
      setLocaleColumns(parsedColumns);
      setRows(cloneRows(builtRows));
      setBaselineRows(cloneRows(builtRows));
      setSourceLoaded(true);
      setValidated(false);
      setExported(false);
      setBuilt(false);
      setLastExport(null);
      setStatusMessage(`Loaded ${builtRows.length} rows from ${file.name}.`);
      return true;
    },
    [],
  );

  const onLoadSource = useCallback(async () => {
    try {
      if (scope === "rebuild") {
        if (!mergedXlsxFile) {
          setStatusMessage("Select merged CSV first.");
          return;
        }
        const ok = await loadFromXlsx(mergedXlsxFile);
        if (ok) {
          setImported(true);
        }
        return;
      }

      if (importMode === "tlk") {
        await loadFromTlk();
        return;
      }
      if (importMode === "xlsx") {
        if (!sourceXlsxFile) {
          setStatusMessage("Select CSV source first.");
          return;
        }
        await loadFromXlsx(sourceXlsxFile);
        return;
      }
      setStatusMessage("Repository branch mode is UI-only in static MVP.");
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : "Source load failed.");
    }
  }, [importMode, loadFromTlk, loadFromXlsx, mergedXlsxFile, scope, sourceXlsxFile]);

  const onParseAndGoEdit = useCallback(() => {
    if (!sourceLoaded) {
      setStatusMessage("Load source first.");
      return;
    }
    if (rows.length === 0 || localeColumns.length === 0) {
      setStatusMessage("No parsed rows/columns to edit. Load source again.");
      return;
    }
    setStepIndex(1);
    setStatusMessage("Schema validated. Edit step is ready.");
  }, [localeColumns.length, rows.length, sourceLoaded]);

  const onCancelExport = useCallback(() => {
    if (!isExportingXlsx) return;
    exportCancelRef.current?.();
    exportCancelRef.current = null;
    exportWorkerRef.current = null;
    setIsExportingXlsx(false);
    setExportProgress(0);
    setExportProgressLabel("");
    setStatusMessage("Export canceled.");
  }, [isExportingXlsx]);

  const runCsvExport = useCallback(async (purpose: "generate" | "publish") => {
    const fileName = csvExportFileName || makeCsvFileName();
    setExportRunMode(purpose);
    setIsExportingXlsx(true);
    setExportProgress(0);
    setExportProgressLabel("0%");
    setStatusMessage(
      purpose === "generate"
        ? "Generating CSV... please wait."
        : "Preparing CSV for PR... please wait.",
    );

    try {
      await yieldToUiFrame();
      const task = createCsvExportTask(rows, localeColumns, fileName, ({ doneRows, totalRows, progress }) => {
        setExportProgress(progress);
        setExportProgressLabel(`${progress}% (${doneRows}/${totalRows})`);
      });
      exportWorkerRef.current = task.worker;
      exportCancelRef.current = task.cancel;
      const exportedCsv = await task.promise;
      exportWorkerRef.current = null;
      exportCancelRef.current = null;
      await saveBytesToDisk(exportedCsv.bytes, exportedCsv.fileName, "text/csv;charset=utf-8");
      setLastExport(exportedCsv);
      setExported(true);
      setExportProgress(100);
      setExportProgressLabel("100%");
      return exportedCsv;
    } finally {
      if (exportWorkerRef.current) {
        exportWorkerRef.current.terminate();
        exportWorkerRef.current = null;
      }
      exportCancelRef.current = null;
      setIsExportingXlsx(false);
      setExportProgress(0);
      setExportProgressLabel("");
    }
  }, [csvExportFileName, localeColumns, rows]);

  const onExportXlsx = useCallback(async () => {
    if (!validated) {
      setStatusMessage("Validate grid before export.");
      return;
    }
    if (isExportingXlsx) {
      return;
    }
    try {
      const exportedCsv = await runCsvExport("generate");
      if (!exportedCsv) return;
      setStatusMessage(`CSV generated: ${exportedCsv.fileName}. Ready to publish.`);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : "CSV export failed.");
    }
  }, [isExportingXlsx, runCsvExport, validated]);

  const onOpenPullRequest = useCallback(async () => {
    if (!validated) {
      setStatusMessage("Validate grid before publish.");
      return;
    }
    if (isExportingXlsx) {
      setStatusMessage("CSV generation is still in progress.");
      return;
    }

    try {
      if (!exported || !lastExport) {
        const exportedCsv = await runCsvExport("publish");
        if (!exportedCsv) return;
      }
      const generated = `https://github.com/nwn-localization/tlk-community-sheet/pull/${Math.floor(Math.random() * 900) + 100}`;
      setPrUrl(generated);
      setStatusMessage("Pull request request accepted (mock 200).");
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : "PR publish failed.");
    }
  }, [exported, isExportingXlsx, lastExport, runCsvExport, validated]);

  const onGenerateDiff = useCallback(() => {
    const diff = computeDiff({
      baselineRows,
      currentRows: rows,
      localeColumns,
      previewLimit: 12,
    });
    setDiffReport(diff.report);
    setStatusMessage(
      `Diff generated: changed=${diff.counts.changed}, conflicts=${diff.counts.conflicts}.`,
    );
  }, [baselineRows, localeColumns, rows]);

  const onDownloadValidationSummary = useCallback(async () => {
    if (!validationFeedback) {
      setStatusMessage("Run validation first.");
      return;
    }

    const missingByColumn = localeColumns.map((col) => {
      let missing = 0;
      for (let i = 0; i < rows.length; i += 1) {
        const value = String(rows[i][col.field] || "").trim();
        if (!value) missing += 1;
      }
      return { title: col.title, missing };
    });

    const validatedCount = Math.max(0, rows.length - validationFeedback.warningCount - validationFeedback.errorCount);
    const summaryLines: string[] = [
      "# TLK Validation Summary",
      `Generated: ${new Date().toISOString()}`,
      `Result level: ${validationFeedback.level.toUpperCase()}`,
      `Can proceed to next step: ${validationFeedback.canProceed ? "YES" : "NO (blocking errors present)"}`,
      "",
      "## Totals",
      `- Total rows: ${rows.length}`,
      `- Validated rows: ${validatedCount}`,
      `- QA warnings (Needs QA): ${validationFeedback.warningCount}`,
      `- Blocking errors: ${validationFeedback.errorCount}`,
      `- Coverage: ${Math.round(coverage * 10) / 10}%`,
      "",
      "## Missing Translations By Column",
      ...missingByColumn.map((item) => `- ${item.title}: ${item.missing}`),
      "",
      "## Warning Samples (up to 20)",
      ...(validationFeedback.warningSamples.length > 0
        ? validationFeedback.warningSamples.map((item) => `- StrRef ${item.strRef}: ${item.reason}`)
        : ["- none"]),
      "",
      "## Blocking Error Samples (up to 20)",
      ...(validationFeedback.errorSamples.length > 0
        ? validationFeedback.errorSamples.map((item) => `- StrRef ${item.strRef}: ${item.reason}`)
        : ["- none"]),
      "",
      "## Rule",
      "- Warnings do not block progression. Only blocking errors lock the next step.",
    ];

    const fileName = `validation_summary_${new Date().toISOString().replace(/[:.]/g, "-")}.md`;
    await saveBytesToDisk(new TextEncoder().encode(summaryLines.join("\n")), fileName, "text/markdown");
    setStatusMessage(`Validation summary downloaded: ${fileName}`);
  }, [coverage, localeColumns, rows, validationFeedback]);

  const onExportDiffJson = useCallback(async () => {
    if (!diffReport) {
      setStatusMessage("Generate diff first.");
      return;
    }
    const content = new TextEncoder().encode(JSON.stringify(diffReport, null, 2));
    await saveBytesToDisk(content, "diff_report.json", "application/json");
    setStatusMessage("Diff JSON exported.");
  }, [diffReport]);

  const onExportDiffMarkdown = useCallback(async () => {
    if (!diffReport) {
      setStatusMessage("Generate diff first.");
      return;
    }
    const md = renderDiffMarkdown(diffReport);
    await saveBytesToDisk(new TextEncoder().encode(md), "diff_report.md", "text/markdown");
    setStatusMessage("Diff Markdown exported.");
  }, [diffReport]);

  const onRebuildTlk = useCallback(async () => {
    if (!rows.length || localeColumns.length === 0) {
      setStatusMessage("Import CSV first.");
      return;
    }

    const generatedArtifacts: ArtifactRow[] = [];
    for (let i = 0; i < localeColumns.length; i += 1) {
      const col = localeColumns[i];
      const locale = String(col.locale || "").toUpperCase().trim();
      const languageId = TLK_LOCALE_TO_LANGUAGE_ID[locale] ?? 0;
      const binary = buildSingleTlkBinaryFromColumn(rows, col.field, languageId);
      const fileName = makeTlkFileName("rebuilt", locale, col.variant === "dialogf");
      await saveBytesToDisk(binary, fileName, "application/octet-stream");
      generatedArtifacts.push({
        file: fileName,
        checksum: quickChecksumHex(binary),
      });
    }

    setArtifacts(generatedArtifacts);
    setBuilt(true);
    setStatusMessage(`Rebuild completed. Generated ${generatedArtifacts.length} TLK artifact(s).`);
  }, [localeColumns, rows]);

  const onUndo = useCallback(() => {
    if (!gridApi) return;
    gridApi.undoCellEditing();
  }, [gridApi]);

  const onRedo = useCallback(() => {
    if (!gridApi) return;
    gridApi.redoCellEditing();
  }, [gridApi]);

  const currentPageSummary = useMemo(() => {
    return rows.length > 0 ? `Rows: ${rows.length}` : "Rows: 0";
  }, [rows.length]);

  const activeStepIndex = useMemo(() => {
    const max = Math.max(0, steps.length - 1);
    return Math.min(Math.max(stepIndex, 0), max);
  }, [stepIndex, steps.length]);

  const stepTitles = steps.map((step) => step.title);
  const stepStates = useMemo(
    () =>
      computeStepUiStates(scope, activeStepIndex, {
        sourceLoaded,
        validated,
        exported,
        imported,
      }),
    [activeStepIndex, exported, imported, scope, sourceLoaded, validated],
  );

  const canContinueExchange =
    Boolean(validationFeedback?.canProceed) &&
    sourceLoaded &&
    rows.length > 0 &&
    localeColumns.length > 0;

  const onStepSelect = useCallback(
    (index: number) => {
      const state = stepStates[index];
      if (!state) return;
      if (!state.canNavigate) {
        setStatusMessage(state.reason || "This step is locked.");
        return;
      }
      setStepIndex(index);
    },
    [stepStates],
  );

  return (
    <main className="workflow-page">
      <header className="workflow-page__topbar">
        <div>
          <h1>TLK Forge</h1>
          <p>Aurora-first workflow: import, edit, publish, rebuild.</p>
        </div>
        <ScopeSwitcher scope={scope} onScopeChange={onScopeChange} />
      </header>

      <StepStrip steps={stepTitles} activeStep={activeStepIndex} stepStates={stepStates} onStepSelect={onStepSelect} />

      <section className="workflow-kpis">
        <article className="kpi">
          <b>{`${activeStepIndex + 1}/${steps.length}`}</b>
          <small>Stage</small>
        </article>
        <article className="kpi">
          <b>{`${Math.round(coverage * 10) / 10}%`}</b>
          <small>Coverage</small>
        </article>
        <article className="kpi">
          <b>{qaWarnings}</b>
          <small>QA warnings</small>
        </article>
        <article className="kpi">
          <b>{blockingErrors}</b>
          <small>Blocking errors</small>
        </article>
        <article className="kpi">
          <b>{rows.length - blockingErrors - qaWarnings}</b>
          <small>Validated rows</small>
        </article>
        <article className="kpi">
          <b>{currentPageSummary}</b>
          <small>Dataset</small>
        </article>
      </section>

      {scope === "exchange" && activeStepIndex === 0 && (
        <section className="workflow-screen">
          <header className="workflow-screen__header">
            <h2>1) Import Source</h2>
            <p>Choose source mode and load TLK/CSV before parsing.</p>
          </header>

          <div className="mode-tabs">
            <button
              type="button"
              className={importMode === "tlk" ? "mode-tabs__active" : ""}
              onClick={() => setImportMode("tlk")}
            >
              From TLK File
            </button>
            <button
              type="button"
              className={importMode === "xlsx" ? "mode-tabs__active" : ""}
              onClick={() => setImportMode("xlsx")}
            >
              From CSV Model
            </button>
            <button
              type="button"
              className={importMode === "repo" ? "mode-tabs__active" : ""}
              onClick={() => setImportMode("repo")}
            >
              From Repository Branch
            </button>
          </div>

          {importMode === "tlk" && (
            <>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>Locale</th>
                      <th>dialog.tlk</th>
                      <th>dialogf.tlk</th>
                      <th>Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tlkBundles.map((bundle) => {
                      const slot = bundleFiles[bundleIdKey(bundle.id)] ?? { dialog: null, dialogf: null };
                      const dialogDropKey = getBundleDropKey(bundle.id, "dialog");
                      const dialogfDropKey = getBundleDropKey(bundle.id, "dialogf");
                      return (
                        <tr key={String(bundle.id)}>
                          <td>
                            <select
                              value={bundle.locale}
                              onChange={(event) =>
                                updateBundle(bundle.id, { locale: event.target.value, dialogfAuto: false })
                              }
                            >
                              {TLK_LOCALE_OPTIONS.map((locale) => (
                                <option key={locale} value={locale}>
                                  {locale}
                                </option>
                              ))}
                            </select>
                          </td>
                          <td>
                            <label
                              className={`bundle-file-picker${activeDropKey === dialogDropKey ? " bundle-file-picker--drag-over" : ""}`}
                              onDragOver={(event) => onBundleDragOver(event, bundle.id, "dialog")}
                              onDragLeave={(event) => onBundleDragLeave(event, bundle.id, "dialog")}
                              onDrop={(event) => onBundleDrop(event, bundle.id, "dialog")}
                            >
                              <span className="bundle-file-picker__button">Choose file</span>
                              <span className="bundle-file-picker__name">{slot.dialog ? slot.dialog.name : "No file selected"}</span>
                              <input
                                type="file"
                                accept=".tlk"
                                className="bundle-file-input"
                                ref={(node) => setBundleInputRef(bundle.id, "dialog", node)}
                                onDragOver={(event) => onBundleDragOver(event, bundle.id, "dialog")}
                                onDragLeave={(event) => onBundleDragLeave(event, bundle.id, "dialog")}
                                onDrop={(event) => onBundleDrop(event, bundle.id, "dialog")}
                                onChange={(event) => updateBundleFile(bundle.id, "dialog", event.target.files?.[0] || null)}
                              />
                            </label>
                            <small className="bundle-file-meta">{slot.dialog ? `Selected: ${slot.dialog.name}` : "Required."}</small>
                          </td>
                          <td>
                            <label
                              className={`bundle-file-picker${activeDropKey === dialogfDropKey ? " bundle-file-picker--drag-over" : ""}`}
                              onDragOver={(event) => onBundleDragOver(event, bundle.id, "dialogf")}
                              onDragLeave={(event) => onBundleDragLeave(event, bundle.id, "dialogf")}
                              onDrop={(event) => onBundleDrop(event, bundle.id, "dialogf")}
                            >
                              <span className="bundle-file-picker__button">Choose file</span>
                              <span className="bundle-file-picker__name">{slot.dialogf ? slot.dialogf.name : "No file selected"}</span>
                              <input
                                type="file"
                                accept=".tlk"
                                className="bundle-file-input"
                                ref={(node) => setBundleInputRef(bundle.id, "dialogf", node)}
                                onDragOver={(event) => onBundleDragOver(event, bundle.id, "dialogf")}
                                onDragLeave={(event) => onBundleDragLeave(event, bundle.id, "dialogf")}
                                onDrop={(event) => onBundleDrop(event, bundle.id, "dialogf")}
                                onChange={(event) => updateBundleFile(bundle.id, "dialogf", event.target.files?.[0] || null)}
                              />
                            </label>
                            <small className="bundle-file-meta">
                              {slot.dialogf ? `Selected: ${slot.dialogf.name}` : "Optional (fallback to dialog.tlk)."}
                            </small>
                          </td>
                          <td>
                            <button type="button" onClick={() => removeBundle(bundle.id)}>
                              Remove
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <div className="table-wrap__actions">
                  <button type="button" onClick={addBundle}>
                    Add Locale Pack
                  </button>
                </div>
              </div>
            </>
          )}

          {importMode === "xlsx" && (
            <div className="workflow-screen__field">
              <label htmlFor="xlsx-source">CSV source</label>
              <input
                id="xlsx-source"
                type="file"
                accept=".csv,.xlsx,.xlsm,.xls"
                onChange={(event) => setSourceXlsxFile(event.target.files?.[0] || null)}
              />
            </div>
          )}

          {importMode === "repo" && (
            <p className="workflow-screen__hint">
              Repo branch import is UI-only in static hosting mode. Data fetch requires dedicated backend/connector.
            </p>
          )}

          <div className="workflow-actions">
            <button type="button" onClick={onLoadSource}>
              Load Source
            </button>
            <button type="button" className="workflow-actions__primary" onClick={onParseAndGoEdit} disabled={!sourceLoaded}>
              Parse &amp; Validate
            </button>
          </div>
        </section>
      )}

      {((scope === "exchange" && activeStepIndex === 1) || (scope === "rebuild" && activeStepIndex === 1)) && (
        <section className="workflow-screen">
          <header className="workflow-screen__header">
            <h2>2) Edit TLK Entries</h2>
            <p>Edit locale columns, then validate before publish/rebuild.</p>
          </header>
          <div className="workflow-actions workflow-actions--editor-tools">
            <button
              type="button"
              className="workflow-actions__primary"
              onClick={applyQaValidation}
              disabled={!sourceLoaded || rows.length === 0 || localeColumns.length === 0}
            >
              Validate
            </button>
            {scope === "exchange" && (
              <button
                type="button"
                className="workflow-actions__secondary workflow-actions__continue"
                onClick={() => setStepIndex(2)}
                disabled={!canContinueExchange}
              >
                Continue to Publish
              </button>
            )}
          </div>
          {validationFeedback ? (
            <div className={`workflow-validation-feedback workflow-validation-feedback--${validationFeedback.level}`}>
              <p>{validationFeedback.message}</p>
              <p className="workflow-validation-feedback__rule">
                Warnings do not block the next step. Only blocking errors lock progression.
              </p>
              <div className="workflow-validation-feedback__actions">
                <button type="button" className="workflow-actions__secondary" onClick={onDownloadValidationSummary}>
                  Download Validation Summary
                </button>
              </div>
            </div>
          ) : null}
          {scope === "exchange" && (
            <div className="locale-config">
              <h4>Optional extra locales for CSV only</h4>
              <div className="locale-row">
                <select
                  onChange={(event) => {
                    const value = event.target.value;
                    if (!value) return;
                    addExtraLocale(value);
                    event.currentTarget.selectedIndex = 0;
                  }}
                  defaultValue=""
                >
                  <option value="" disabled>
                    Select locale...
                  </option>
                  {availableExtraLocales.map((locale) => (
                    <option key={locale} value={locale}>
                      {locale}
                    </option>
                  ))}
                </select>
              </div>
              <div className="locale-chip-list">
                {normalizedExtraLocales.map((locale) => (
                  <span key={locale} className="locale-chip">
                    {locale}
                    <button type="button" onClick={() => removeExtraLocale(locale)}>
                      ×
                    </button>
                  </span>
                ))}
              </div>
            </div>
          )}
          {rows.length === 0 || localeColumns.length === 0 ? (
            <div className="workflow-empty-state">
              <p className="workflow-screen__hint">No parsed rows to show yet. Go back to Import and load TLK/CSV first.</p>
              <div className="workflow-actions">
                <button type="button" className="workflow-actions__secondary" onClick={() => setStepIndex(0)}>
                  Back to Import
                </button>
              </div>
            </div>
          ) : (
            <LocalizationGrid
              rows={rows}
              localeColumns={localeColumns}
              pageSize={pageSize}
              onPageSizeChange={setPageSize}
              onRowsChange={onGridRowsChange}
              onGridApiReady={onGridApiReady}
              onUndo={onUndo}
              onRedo={onRedo}
            />
          )}
        </section>
      )}

      {scope === "exchange" && activeStepIndex === 2 && (
        <section className="workflow-screen">
          <header className="workflow-screen__header">
            <h2>3) Publish</h2>
          </header>
          <div className="grid-2">
            <article className="card">
              <h3>GitHub PR</h3>
              <div className="workflow-actions">
                <button type="button" className="workflow-actions__primary" onClick={onOpenPullRequest} disabled={!validated || isExportingXlsx}>
                  Open PR
                </button>
              </div>
              {prUrl && (
                <p className="workflow-screen__hint">
                  PR URL: <a href={prUrl} target="_blank" rel="noreferrer">{prUrl}</a>
                </p>
              )}
            </article>
            <article className="card">
              <h3>Export CSV</h3>
              <p className="workflow-screen__hint">{`Target CSV: ${csvExportFileName}`}</p>
              <div className="workflow-actions">
                <button type="button" className="workflow-actions__primary" onClick={onExportXlsx} disabled={!validated || isExportingXlsx}>
                  {isExportingXlsx && exportRunMode === "generate" ? "Generating CSV..." : "Generate CSV"}
                </button>
                {isExportingXlsx ? (
                  <button type="button" className="workflow-actions__secondary" onClick={onCancelExport}>
                    Cancel
                  </button>
                ) : null}
              </div>
              {isExportingXlsx ? (
                <div className="workflow-inline-progress" role="status" aria-live="polite">
                  <span className="workflow-inline-progress__label">
                    <span className="workflow-inline-progress__spinner" aria-hidden="true" />
                    {`${
                      exportRunMode === "publish" ? "Preparing CSV for PR..." : "Generating CSV..."
                    } ${exportProgressLabel || `${exportProgress}%`}`}
                  </span>
                  <div className="workflow-inline-progress__bar" aria-hidden="true">
                    <span style={{ width: `${Math.max(0, Math.min(100, exportProgress))}%` }} />
                  </div>
                </div>
              ) : null}
              {lastExport ? <p className="workflow-screen__hint">{`Last CSV: ${lastExport.fileName}`}</p> : null}
            </article>
          </div>
        </section>
      )}

      {scope === "rebuild" && activeStepIndex === 0 && (
        <section className="workflow-screen">
          <header className="workflow-screen__header">
            <h2>Rebuild Scope: 1) Load Approved CSV</h2>
            <p>Import reviewed CSV and prepare rebuild.</p>
          </header>
          <div className="workflow-screen__field">
            <label htmlFor="merged-xlsx">Merged CSV</label>
            <input
              id="merged-xlsx"
              type="file"
              accept=".csv,.xlsx,.xlsm,.xls"
              onChange={(event) => setMergedXlsxFile(event.target.files?.[0] || null)}
            />
          </div>
          <div className="workflow-actions">
            <button type="button" className="workflow-actions__primary" onClick={onLoadSource}>
              Import Merged CSV
            </button>
            <button type="button" onClick={onGenerateDiff} disabled={!sourceLoaded}>
              Generate Diff Report
            </button>
          </div>
          {diffReport && (
            <div className="diff-report">
              <p>{`Changed: ${diffReport.summary.changed}, Added: ${diffReport.summary.added}, Removed: ${diffReport.summary.removed}, Conflicts: ${diffReport.summary.conflicts}`}</p>
              <div className="workflow-actions">
                <button type="button" onClick={onExportDiffJson}>
                  Export Diff JSON
                </button>
                <button type="button" onClick={onExportDiffMarkdown}>
                  Export Diff Markdown
                </button>
                <button type="button" className="workflow-actions__secondary" onClick={() => setStepIndex(1)} disabled={!imported}>
                  Continue to Rebuild
                </button>
              </div>
            </div>
          )}
        </section>
      )}

      {scope === "rebuild" && activeStepIndex === 1 && (
        <section className="workflow-screen">
          <header className="workflow-screen__header">
            <h2>Rebuild Scope: 2) Rebuild TLK/TLKF</h2>
            <p>Generate TLK artifacts from loaded CSV dataset.</p>
          </header>
          <div className="workflow-actions">
            <button type="button" className="workflow-actions__primary" onClick={onRebuildTlk} disabled={!sourceLoaded}>
              Build TLK
            </button>
          </div>
          {built && (
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>File</th>
                    <th>Status</th>
                    <th>Checksum</th>
                  </tr>
                </thead>
                <tbody>
                  {artifacts.map((artifact) => (
                    <tr key={artifact.file}>
                      <td>{artifact.file}</td>
                      <td>
                        <span className={statusBadgeClass("validated")}>Built</span>
                      </td>
                      <td>{artifact.checksum}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </section>
      )}
    </main>
  );
};

export default LocalizationWorkflow;




