import { useCallback, useMemo, useRef, useState } from "react";
import type { GridApi } from "ag-grid-community";
import * as XLSX from "xlsx";

import ScopeSwitcher from "../../components/ScopeSwitcher";
import StepStrip from "../../components/StepStrip";
import LocalizationGrid from "../../components/LocalizationGrid";
import {
  buildRowsFromParsedTlkBundles,
  buildSingleTlkBinaryFromColumn,
  parseSingleTlkBuffer,
  quickChecksumHex,
  safeFileNameFromPath,
  tlkEncodingForLanguageId,
  TLK_LOCALE_TO_LANGUAGE_ID,
} from "../../lib/tlk";
import { localeCodeToFieldToken, normalizeLocaleCode, type LocaleColumn, type ParsedTlkBundle, type TlkBundleConfig, type TlkGridRow } from "../../lib/types";
import { applyDialogfFallbacks, validateTlkBundles } from "../../lib/validation";
import {
  buildTlkRowsFromXlsxRows,
  computeCsvExportChangeSummary,
  getParsedLocaleColumnsFromRows,
  makeCsvFileName,
  parseWorkbookFile,
} from "../../lib/xlsx";
import {
  readSessionValue,
  removeSessionValue,
  SESSION_KEY_GITHUB_BASE_BRANCH,
  SESSION_KEY_GITHUB_CSV_FOLDER,
  SESSION_KEY_GITHUB_REPO,
  SESSION_KEY_GITHUB_TOKEN,
  writeSessionValue,
} from "../../lib/sessionStorage";
import { fetchRepoFileFromGitHub, publishCsvToGitHub } from "../../lib/github";
import { buildZipArchive } from "../../lib/zip";
import { computeStepUiStates } from "../../lib/workflowProgress";
import type { ImportMode, WorkflowScope, WorkflowStep } from "./types";

type ArtifactRow = {
  file: string;
  checksum: string;
  bytes: Uint8Array;
  locale: string;
  encoding: string;
  isDialogf: boolean;
  kind: "tlk" | "zip";
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

type CsvEncoding =
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

const CSV_ENCODING_OPTIONS: Array<{ value: CsvEncoding; label: string }> = [
  { value: "utf8-bom", label: "UTF-8 (BOM)" },
  { value: "utf8", label: "UTF-8 (no BOM)" },
  { value: "iso-8859-1", label: "ISO-8859-1" },
  { value: "windows-1252", label: "Windows-1252" },
  { value: "windows-1251", label: "Windows-1251" },
  { value: "euc-jp", label: "EUC-JP" },
  { value: "shift_jis", label: "Shift_JIS" },
  { value: "euc-kr", label: "EUC-KR" },
  { value: "windows-1250", label: "Windows-1250" },
  { value: "gb2312", label: "GB2312" },
  { value: "iso-8859-2", label: "ISO-8859-2" },
];

const TLK_LOCALE_OPTIONS = ["EN", "PL", "DE", "FR", "ES", "IT", "PT-BR", "RU"];
const DEFAULT_GITHUB_REPO = "enonwow/nwn-localization-test";
const DEFAULT_GITHUB_BASE_BRANCH = "main";
const DEFAULT_GITHUB_CSV_FOLDER = "csv-latest";
const DEFAULT_CSV_FILE_NAME = "test.csv";

function normalizeGithubRepoRef(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";

  if (/^https?:\/\//i.test(raw)) {
    try {
      const url = new URL(raw);
      const parts = url.pathname.split("/").filter(Boolean);
      if (parts.length >= 2) {
        return `${parts[0]}/${parts[1].replace(/\.git$/i, "")}`;
      }
      return "";
    } catch {
      return "";
    }
  }

  const normalized = raw
    .replace(/^github\.com\//i, "")
    .replace(/\.git$/i, "")
    .replace(/^\/+|\/+$/g, "");
  const parts = normalized.split("/").filter(Boolean);
  if (parts.length < 2) return "";
  return `${parts[0]}/${parts[1]}`;
}

function normalizeGithubPath(value: string, fallback: string): string {
  const normalized = String(value || "").trim().replace(/^\/+|\/+$/g, "");
  return normalized || fallback;
}

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
  lineEnding: "lf" | "crlf",
  encoding: CsvEncoding,
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

  worker.postMessage({ rows, localeColumns, fileName, lineEnding, encoding });
  const cancel = () => {
    if (isSettled) return;
    isSettled = true;
    worker.terminate();
    rejectPromise?.(new Error("Export canceled."));
  };
  return { worker, promise, cancel };
}

function csvEncodingToMimeCharset(encoding: CsvEncoding): string {
  switch (encoding) {
    case "utf8":
    case "utf8-bom":
      return "utf-8";
    case "iso-8859-1":
      return "iso-8859-1";
    case "windows-1252":
      return "windows-1252";
    case "windows-1251":
      return "windows-1251";
    case "euc-jp":
      return "euc-jp";
    case "shift_jis":
      return "shift_jis";
    case "euc-kr":
      return "euc-kr";
    case "windows-1250":
      return "windows-1250";
    case "gb2312":
      return "gb2312";
    case "iso-8859-2":
      return "iso-8859-2";
    default:
      return "utf-8";
  }
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

function isCsvLikeFileName(name: string | undefined): boolean {
  return /\.csv$/i.test(String(name || "").trim());
}

function isSpreadsheetSourceFile(file: File): boolean {
  return /\.(csv|xlsx|xlsm|xls)$/i.test(String(file?.name || "").trim());
}

function makeSafeDownloadStem(rawValue: string, fallback: string): string {
  const trimmed = String(rawValue || "").trim();
  if (!trimmed) return fallback;
  const sanitized = trimmed
    .replace(/\.[^.]+$/u, "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/gu, "_")
    .trim();
  return sanitized || fallback;
}

function makeRebuildZipName(sourceName: string): string {
  const fileName = safeFileNameFromPath(sourceName);
  const stem = makeSafeDownloadStem(fileName, "localization");
  return `${stem}.zip`;
}

function makeLocaleFileToken(locale: string): string {
  const normalized = String(locale || "").toUpperCase().trim();
  const token = normalized.replace(/[^A-Z0-9]+/g, "-");
  return token || "XX";
}

function makeRebuildArtifactName(baseStem: string, locale: string, isDialogf: boolean): string {
  const localeToken = makeLocaleFileToken(locale);
  const stem = isDialogf ? `${baseStem}f` : baseStem;
  return `${stem}_${localeToken}.tlk`;
}

function makeZipLocaleFolder(locale: string): string {
  return localeCodeToFieldToken(locale).replace(/_/g, "-");
}

function makeZipArtifactPath(rootStem: string, baseStem: string, locale: string, isDialogf: boolean): string {
  const folder = makeZipLocaleFolder(locale);
  const file = isDialogf ? `${baseStem}f.tlk` : `${baseStem}.tlk`;
  return `${rootStem}/${folder}/${file}`;
}

async function detectCsvLineEnding(file: File): Promise<"lf" | "crlf"> {
  if (!isCsvLikeFileName(file?.name)) {
    return "lf";
  }
  try {
    const probeSize = Math.min(file.size, 1024 * 1024);
    const probeBuffer = await file.slice(0, probeSize).arrayBuffer();
    const probeText = new TextDecoder().decode(probeBuffer);
    return probeText.includes("\r\n") ? "crlf" : "lf";
  } catch {
    return "lf";
  }
}

const BATCH_HISTORY_LIMIT = 25;

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
  const [githubRepoInput, setGithubRepoInput] = useState(() =>
    readSessionValue(SESSION_KEY_GITHUB_REPO, DEFAULT_GITHUB_REPO),
  );
  const [githubBaseBranchInput, setGithubBaseBranchInput] = useState(() =>
    readSessionValue(SESSION_KEY_GITHUB_BASE_BRANCH, DEFAULT_GITHUB_BASE_BRANCH),
  );
  const [githubCsvFolderInput, setGithubCsvFolderInput] = useState(() =>
    readSessionValue(SESSION_KEY_GITHUB_CSV_FOLDER, DEFAULT_GITHUB_CSV_FOLDER),
  );
  const [repoImportBranch, setRepoImportBranch] = useState(DEFAULT_GITHUB_BASE_BRANCH);
  const [repoImportPath, setRepoImportPath] = useState(`${DEFAULT_GITHUB_CSV_FOLDER}/${DEFAULT_CSV_FILE_NAME}`);
  const [rows, setRows] = useState<TlkGridRow[]>([]);
  const [baselineRows, setBaselineRows] = useState<TlkGridRow[]>([]);
  const [batchUndoStack, setBatchUndoStack] = useState<TlkGridRow[][]>([]);
  const [batchRedoStack, setBatchRedoStack] = useState<TlkGridRow[][]>([]);
  const [localeColumns, setLocaleColumns] = useState<LocaleColumn[]>([]);
  const [sourceLoaded, setSourceLoaded] = useState(false);
  const [validated, setValidated] = useState(false);
  const [exported, setExported] = useState(false);
  const [imported, setImported] = useState(false);
  const [built, setBuilt] = useState(false);
  const [artifacts, setArtifacts] = useState<ArtifactRow[]>([]);
  const [pageSize, setPageSize] = useState(25);
  const [, setStatusMessage] = useState("Ready. Aurora-first workflow is active.");
  const [validationFeedback, setValidationFeedback] = useState<ValidationFeedback | null>(null);
  const [isLoadingSource, setIsLoadingSource] = useState(false);
  const [loadSourceLabel, setLoadSourceLabel] = useState("");
  const [loadSourceError, setLoadSourceError] = useState("");
  const [isExportingXlsx, setIsExportingXlsx] = useState(false);
  const [isPublishingPr, setIsPublishingPr] = useState(false);
  const [exportProgress, setExportProgress] = useState(0);
  const [exportProgressLabel, setExportProgressLabel] = useState("");
  const [lastExport, setLastExport] = useState<{ fileName: string; bytes: Uint8Array } | null>(null);
  const [exportRunMode, setExportRunMode] = useState<"generate" | "publish">("generate");
  const [csvExportFileName, setCsvExportFileName] = useState(DEFAULT_CSV_FILE_NAME);
  const [csvLineEnding, setCsvLineEnding] = useState<"lf" | "crlf">("lf");
  const [csvEncoding, setCsvEncoding] = useState<CsvEncoding>("utf8-bom");
  const [activeDropKey, setActiveDropKey] = useState<string | null>(null);
  const [prUrl, setPrUrl] = useState("");
  const [prError, setPrError] = useState("");
  const [githubRuntimeToken, setGithubRuntimeToken] = useState(() => {
    removeSessionValue(SESSION_KEY_GITHUB_TOKEN);
    return "";
  });
  const [gridApi, setGridApi] = useState<GridApi<TlkGridRow> | null>(null);
  const exportWorkerRef = useRef<Worker | null>(null);
  const exportCancelRef = useRef<(() => void) | null>(null);
  const bundleInputRefs = useRef<Record<string, HTMLInputElement | null>>({});
  const githubToken = String(githubRuntimeToken || "").trim();
  const githubRepo = normalizeGithubRepoRef(githubRepoInput);
  const githubBaseBranch = normalizeGithubPath(githubBaseBranchInput, DEFAULT_GITHUB_BASE_BRANCH);
  const githubCsvFolder = normalizeGithubPath(githubCsvFolderInput, DEFAULT_GITHUB_CSV_FOLDER);
  const githubCsvFolderUrl = githubRepo
    ? `https://github.com/${githubRepo}/tree/${githubBaseBranch}/${githubCsvFolder}`
    : "";
  const isGithubRepoValid = githubRepo.length > 0;
  const onGithubTokenChange = useCallback((nextValue: string) => {
    setGithubRuntimeToken(nextValue);
    removeSessionValue(SESSION_KEY_GITHUB_TOKEN);
  }, []);
  const onGithubRepoInputChange = useCallback((nextValue: string) => {
    setGithubRepoInput(nextValue);
    writeSessionValue(SESSION_KEY_GITHUB_REPO, nextValue);
  }, []);
  const onGithubBaseBranchInputChange = useCallback((nextValue: string) => {
    setGithubBaseBranchInput(nextValue);
    writeSessionValue(SESSION_KEY_GITHUB_BASE_BRANCH, nextValue);
  }, []);
  const onGithubCsvFolderInputChange = useCallback((nextValue: string) => {
    setGithubCsvFolderInput(nextValue);
    writeSessionValue(SESSION_KEY_GITHUB_CSV_FOLDER, nextValue);
  }, []);

  const onCsvPickerDragOver = useCallback((event: React.DragEvent<HTMLElement>, dropKey: string) => {
    event.preventDefault();
    event.stopPropagation();
    if (activeDropKey !== dropKey) {
      setActiveDropKey(dropKey);
    }
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = "copy";
    }
  }, [activeDropKey]);

  const onCsvPickerDragLeave = useCallback((event: React.DragEvent<HTMLElement>, dropKey: string) => {
    event.preventDefault();
    event.stopPropagation();
    setActiveDropKey((prev) => (prev === dropKey ? null : prev));
  }, []);

  const onCsvPickerDrop = useCallback((
    event: React.DragEvent<HTMLElement>,
    dropKey: string,
    onFileChange: (file: File | null) => void,
  ) => {
    event.preventDefault();
    event.stopPropagation();
    setActiveDropKey((prev) => (prev === dropKey ? null : prev));

    const files = event.dataTransfer?.files;
    if (!files || files.length === 0) {
      setStatusMessage("No file dropped.");
      return;
    }

    const file = files[0];
    if (!isSpreadsheetSourceFile(file)) {
      setStatusMessage("Drop a CSV/XLSX file.");
      return;
    }

    onFileChange(file);
    if (files.length > 1) {
      setStatusMessage(`Multiple files dropped, using first: ${file.name}`);
      return;
    }
    setStatusMessage(`Loaded from drop: ${file.name}`);
  }, []);

  const renderRepoSourceFields = useCallback(() => (
    <>
      <div className="workflow-screen__field">
        <label htmlFor="repo-name">Repository (owner/name or URL)</label>
        <input
          id="repo-name"
          type="text"
          value={githubRepoInput}
          onChange={(event) => onGithubRepoInputChange(event.target.value)}
          placeholder={DEFAULT_GITHUB_REPO}
        />
      </div>
      <div className="workflow-screen__field">
        <label htmlFor="repo-branch">Repository branch</label>
        <input
          id="repo-branch"
          type="text"
          value={repoImportBranch}
          onChange={(event) => setRepoImportBranch(event.target.value)}
          placeholder={githubBaseBranch}
        />
      </div>
      <div className="workflow-screen__field">
        <label htmlFor="repo-path">Repository CSV path</label>
        <input
          id="repo-path"
          type="text"
          value={repoImportPath}
          onChange={(event) => setRepoImportPath(event.target.value)}
          placeholder={`${githubCsvFolder}/${csvExportFileName}`}
        />
      </div>
      <div className="workflow-screen__field">
        <label htmlFor="repo-token">GitHub PAT (optional for public repo)</label>
        <input
          id="repo-token"
          type="password"
          placeholder="github_pat_..."
          value={githubRuntimeToken}
          onChange={(event) => onGithubTokenChange(event.target.value)}
        />
      </div>
      <p className="workflow-screen__hint">
        {isGithubRepoValid ? (
          <>
            Loads CSV from <code>{githubRepo}</code> using branch + path.
          </>
        ) : (
          "Provide a valid GitHub repository first (owner/name or URL)."
        )}
      </p>
    </>
  ), [
    csvExportFileName,
    githubBaseBranch,
    githubCsvFolder,
    githubRepo,
    githubRepoInput,
    githubRuntimeToken,
    isGithubRepoValid,
    onGithubTokenChange,
    repoImportBranch,
    repoImportPath,
  ]);

  const renderCsvModelFilePicker = useCallback(
    (options: {
      inputId: string;
      selectedFile: File | null;
      onFileChange: (file: File | null) => void;
    }) => {
      const dropKey = `csv:${options.inputId}`;
      return (
      <div className="workflow-screen__field">
        <label
          className={`bundle-file-picker bundle-file-picker--csv${activeDropKey === dropKey ? " bundle-file-picker--drag-over" : ""}`}
          onDragOver={(event) => onCsvPickerDragOver(event, dropKey)}
          onDragLeave={(event) => onCsvPickerDragLeave(event, dropKey)}
          onDrop={(event) => onCsvPickerDrop(event, dropKey, options.onFileChange)}
        >
          <span className="bundle-file-picker__button">Choose file</span>
          <span className="bundle-file-picker__name">
            {options.selectedFile ? options.selectedFile.name : "No file selected"}
          </span>
          <input
            id={options.inputId}
            type="file"
            accept=".csv,.xlsx,.xlsm,.xls"
            className="bundle-file-input"
            onDragOver={(event) => onCsvPickerDragOver(event, dropKey)}
            onDragLeave={(event) => onCsvPickerDragLeave(event, dropKey)}
            onDrop={(event) => onCsvPickerDrop(event, dropKey, options.onFileChange)}
            onChange={(event) => options.onFileChange(event.target.files?.[0] || null)}
          />
        </label>
      </div>
      );
    },
    [activeDropKey, onCsvPickerDragLeave, onCsvPickerDragOver, onCsvPickerDrop],
  );

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
  const emptyRows = useMemo(
    () =>
      rows.filter((row) => {
        const allLocaleValuesEmpty =
          localeColumns.length === 0 ||
          localeColumns.every((col) => String(row[col.field] || "").trim().length === 0);
        return allLocaleValuesEmpty;
      }).length,
    [localeColumns, rows],
  );
  const changeSummary = useMemo(() => {
    if (!sourceLoaded || rows.length === 0 || localeColumns.length === 0 || baselineRows.length === 0) {
      return {
        changedRows: 0,
        changedCells: 0,
        addedRows: 0,
        removedRows: 0,
        sourceChangedRows: 0,
        changedRowStrRefs: [] as number[],
        touchedLocales: [] as string[],
      };
    }
    return computeCsvExportChangeSummary({
      baselineRows,
      currentRows: rows,
      localeColumns,
    });
  }, [baselineRows, localeColumns, rows, sourceLoaded]);

  const hasPublishableChanges = changeSummary.changedRows > 0;
  const clearBatchHistory = useCallback(() => {
    setBatchUndoStack([]);
    setBatchRedoStack([]);
  }, []);

  const resetFlowState = useCallback(() => {
    setRows([]);
    setBaselineRows([]);
    setLocaleColumns([]);
    setSourceLoaded(false);
    setValidated(false);
    setExported(false);
    setImported(false);
    setBuilt(false);
    setArtifacts([]);
    setBundleFiles({});
    setValidationFeedback(null);
    if (exportWorkerRef.current) {
      exportWorkerRef.current.terminate();
      exportWorkerRef.current = null;
    }
    exportCancelRef.current = null;
    setIsExportingXlsx(false);
    setIsPublishingPr(false);
    setExportProgress(0);
    setExportProgressLabel("");
    setLastExport(null);
    setCsvExportFileName(DEFAULT_CSV_FILE_NAME);
    setCsvLineEnding("lf");
    setLoadSourceError("");
    setActiveDropKey(null);
    clearBatchHistory();
  }, [clearBatchHistory]);

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

  const applyRowsAndResetMeta = useCallback((nextRows: TlkGridRow[]) => {
    setRows(nextRows);
    setValidated(false);
    setExported(false);
    setBuilt(false);
    setLastExport(null);
    setValidationFeedback(null);
  }, []);

  const onGridRowsChange = useCallback((nextRows: TlkGridRow[]) => {
    applyRowsAndResetMeta(nextRows);
  }, [applyRowsAndResetMeta]);

  const onGridBatchRowsChange = useCallback((nextRows: TlkGridRow[], previousRows?: TlkGridRow[]) => {
    const snapshotBeforeChange = previousRows ? cloneRows(previousRows) : cloneRows(rows);
    setBatchUndoStack((prev) => [snapshotBeforeChange, ...prev].slice(0, BATCH_HISTORY_LIMIT));
    setBatchRedoStack([]);
    applyRowsAndResetMeta(nextRows);
  }, [applyRowsAndResetMeta, rows]);

  const onAddGridRow = useCallback(() => {
    const previousRows = cloneRows(rows);
    const maxStrRef = previousRows.reduce((max, row) => {
      const value = Number(row.strRef);
      if (!Number.isFinite(value)) return max;
      return Math.max(max, value);
    }, -1);
    const nextStrRef = maxStrRef + 1;

    const nextRow: TlkGridRow = {
      id: nextStrRef,
      strRef: nextStrRef,
      sourceEn: "",
      context: "",
      status: "Draft",
    };

    for (let i = 0; i < localeColumns.length; i += 1) {
      const col = localeColumns[i];
      nextRow[col.field] = "";
    }

    const nextRows = [...previousRows, nextRow];
    setBatchUndoStack((prev) => [previousRows, ...prev].slice(0, BATCH_HISTORY_LIMIT));
    setBatchRedoStack([]);
    applyRowsAndResetMeta(nextRows);
    setStatusMessage(`Added new row StrRef ${nextStrRef}.`);
    return nextStrRef;
  }, [applyRowsAndResetMeta, localeColumns, rows]);

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
              ? `Validation blocked: Missing=${emptyRows}, Errors=${tokenMismatch}, Validated=${validatedCount}.`
              : `Validation warnings: Missing=${emptyRows}, Errors=${tokenMismatch}, Validated=${validatedCount}.`,
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
      const message = validation.issues[0]?.message || "TLK bundle validation failed.";
      setStatusMessage(message);
      setLoadSourceError(message);
      return;
    }

    const fallback = applyDialogfFallbacks(validation.normalized);
    setTlkBundles(fallback.bundles);

    const parsedBundles: ParsedTlkBundle[] = [];
    for (let i = 0; i < fallback.bundles.length; i += 1) {
      const bundle = fallback.bundles[i];
      const dialogFile = getBundleFile(bundle.id, "dialog");
      if (!dialogFile) {
        const message = `Locale ${bundle.locale}: select dialog.tlk file.`;
        setStatusMessage(message);
        setLoadSourceError(message);
        return;
      }
      const parsedDialog = parseSingleTlkBuffer(await dialogFile.arrayBuffer(), dialogFile.name);

      let parsedDialogf = null;
      if (!bundle.dialogfAuto && String(bundle.dialogf || "").trim().length > 0) {
        const dialogfFile = getBundleFile(bundle.id, "dialogf");
        if (!dialogfFile) {
          const message = `Locale ${bundle.locale}: select dialogf.tlk file or use fallback.`;
          setStatusMessage(message);
          setLoadSourceError(message);
          return;
        }
        parsedDialogf = parseSingleTlkBuffer(await dialogfFile.arrayBuffer(), dialogfFile.name);
        if (parsedDialogf.entryCount !== parsedDialog.entryCount) {
          const message = `Entry mismatch for locale ${bundle.locale}: dialog vs dialogf.`;
          setStatusMessage(message);
          setLoadSourceError(message);
          return;
        }
      }
      parsedBundles.push({ bundle, parsed: parsedDialog, dialogfParsed: parsedDialogf });
    }

    const baseCount = parsedBundles[0]?.parsed.entryCount || 0;
    const mismatch = parsedBundles.find((item) => item.parsed.entryCount !== baseCount);
    if (mismatch) {
      const message = "Entry count mismatch across locale packs.";
      setStatusMessage(message);
      setLoadSourceError(message);
      return;
    }

    const effectiveColumns = buildEffectiveLocaleColumns(fallback.bundles, normalizedExtraLocales);
    const builtRows = buildRowsFromParsedTlkBundles(parsedBundles, effectiveColumns);
    const primarySourceName = parsedBundles[0]?.parsed.fileName || fallback.bundles[0]?.dialog || "";
    setCsvExportFileName(makeCsvFileName(primarySourceName));
    setCsvLineEnding("lf");
    setLocaleColumns(effectiveColumns);
    setRows(cloneRows(builtRows));
    setBaselineRows(cloneRows(builtRows));
    clearBatchHistory();
    setSourceLoaded(true);
    setValidated(false);
    setExported(false);
    setBuilt(false);
    setLastExport(null);
    setLoadSourceError("");
    setStatusMessage(
      fallback.fallbackCount > 0
        ? `Loaded ${builtRows.length} rows with dialogf fallback for ${fallback.fallbackCount} pack(s).`
        : `Loaded ${builtRows.length} rows from TLK bundles.`,
    );
  }, [clearBatchHistory, getBundleFile, normalizedExtraLocales, tlkBundles]);

  const loadFromXlsx = useCallback(
    async (file: File) => {
      const workbook = await parseWorkbookFile(file, XLSX);
      const builtRows = buildTlkRowsFromXlsxRows(workbook.rows);
      if (builtRows.length === 0) {
        const message = "CSV has no StrRef rows.";
        setStatusMessage(message);
        setLoadSourceError(message);
        return false;
      }
      const parsedColumns = getParsedLocaleColumnsFromRows(builtRows);
      const parsedExtraLocales = Array.from(
        new Set(
          parsedColumns
            .filter((col) => col.variant === "xlsx-extra")
            .map((col) => normalizeLocaleCode(col.locale)),
        ),
      );
      setCsvExportFileName(makeCsvFileName(file.name));
      setCsvLineEnding(await detectCsvLineEnding(file));
      setExtraXlsxLocales(parsedExtraLocales);
      setLocaleColumns(parsedColumns);
      setRows(cloneRows(builtRows));
      setBaselineRows(cloneRows(builtRows));
      clearBatchHistory();
      setSourceLoaded(true);
      setValidated(false);
      setExported(false);
      setBuilt(false);
      setLastExport(null);
      setLoadSourceError("");
      setStatusMessage(`Loaded ${builtRows.length} rows from ${file.name}.`);
      return true;
    },
    [clearBatchHistory],
  );

  const loadFromRepo = useCallback(async () => {
    if (!isGithubRepoValid) {
      const message = "Set a valid repository first (owner/name or URL).";
      setStatusMessage(message);
      setLoadSourceError(message);
      return false;
    }

    const branch = String(repoImportBranch || "").trim();
    const filePath = String(repoImportPath || "").trim();
    if (!branch) {
      const message = "Set repository branch first.";
      setStatusMessage(message);
      setLoadSourceError(message);
      return false;
    }
    if (!filePath) {
      const message = "Set repository CSV file path first.";
      setStatusMessage(message);
      setLoadSourceError(message);
      return false;
    }

    const fetched = await fetchRepoFileFromGitHub({
      token: githubToken || undefined,
      repoFullName: githubRepo,
      branch,
      filePath,
    });
    const file = new File([fetched.bytes], fetched.fileName, { type: "text/csv;charset=utf-8" });
    const loaded = await loadFromXlsx(file);
    if (loaded) {
      setLoadSourceError("");
      setStatusMessage(`Loaded ${fetched.filePath} from ${githubRepo}@${branch}.`);
    }
    return loaded;
  }, [githubRepo, githubToken, isGithubRepoValid, loadFromXlsx, repoImportBranch, repoImportPath]);

  const onLoadSource = useCallback(async () => {
    if (isLoadingSource) {
      return;
    }

    setLoadSourceError("");
    setIsLoadingSource(true);
    setLoadSourceLabel(
      scope === "rebuild"
        ? "Importing merged CSV..."
        : importMode === "repo"
          ? "Loading CSV from repository..."
          : importMode === "xlsx"
            ? "Loading CSV model..."
            : "Loading TLK bundles...",
    );
    try {
      if (scope === "rebuild") {
        if (importMode === "repo") {
          const ok = await loadFromRepo();
          if (ok) {
            setLoadSourceError("");
            setImported(true);
          }
          return;
        }
        if (importMode === "xlsx") {
          if (!mergedXlsxFile) {
            const message = "Select CSV source first.";
            setStatusMessage(message);
            setLoadSourceError(message);
            return;
          }
          const ok = await loadFromXlsx(mergedXlsxFile);
          if (ok) {
            setLoadSourceError("");
            setImported(true);
          }
          return;
        }
        const message = "Unsupported import mode.";
        setStatusMessage(message);
        setLoadSourceError(message);
        return;
      }

      if (importMode === "tlk") {
        await loadFromTlk();
        return;
      }
      if (importMode === "xlsx") {
        if (!sourceXlsxFile) {
          const message = "Select CSV source first.";
          setStatusMessage(message);
          setLoadSourceError(message);
          return;
        }
        await loadFromXlsx(sourceXlsxFile);
        return;
      }
      if (importMode === "repo") {
        await loadFromRepo();
        return;
      }
      const message = "Unsupported import mode.";
      setStatusMessage(message);
      setLoadSourceError(message);
    } catch (error) {
      const message = error instanceof Error ? error.message : "Source load failed.";
      setStatusMessage(message);
      setLoadSourceError(message);
    } finally {
      setIsLoadingSource(false);
      setLoadSourceLabel("");
    }
  }, [importMode, isLoadingSource, loadFromRepo, loadFromTlk, loadFromXlsx, mergedXlsxFile, scope, sourceXlsxFile]);

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

  const runCsvExport = useCallback(async (
    purpose: "generate" | "publish",
    options?: { saveToDisk?: boolean; lineEnding?: "lf" | "crlf" },
  ) => {
    const shouldSaveToDisk = options?.saveToDisk ?? (purpose === "generate");
    const exportLineEnding = options?.lineEnding ?? csvLineEnding;
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
      const task = createCsvExportTask(rows, localeColumns, fileName, exportLineEnding, csvEncoding, ({ doneRows, totalRows, progress }) => {
        setExportProgress(progress);
        setExportProgressLabel(`${progress}% (${doneRows}/${totalRows})`);
      });
      exportWorkerRef.current = task.worker;
      exportCancelRef.current = task.cancel;
      const exportedCsv = await task.promise;
      exportWorkerRef.current = null;
      exportCancelRef.current = null;
      if (shouldSaveToDisk) {
        const charset = csvEncodingToMimeCharset(csvEncoding);
        await saveBytesToDisk(exportedCsv.bytes, exportedCsv.fileName, `text/csv;charset=${charset}`);
      }
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
  }, [csvEncoding, csvExportFileName, csvLineEnding, localeColumns, rows]);

  const onExportXlsx = useCallback(async () => {
    if (isExportingXlsx) {
      return;
    }
    try {
      const exportedCsv = await runCsvExport("generate", { saveToDisk: true });
      if (!exportedCsv) return;
      setStatusMessage(`CSV generated: ${exportedCsv.fileName}. Ready to publish.`);
    } catch (error) {
      setStatusMessage(error instanceof Error ? error.message : "CSV export failed.");
    }
  }, [isExportingXlsx, runCsvExport]);

  const detectRepoTargetLineEnding = useCallback(
    async (fileName: string): Promise<"lf" | "crlf"> => {
      try {
        const targetPath = `${githubCsvFolder}/${fileName}`;
        const fetched = await fetchRepoFileFromGitHub({
          token: githubToken || undefined,
          repoFullName: githubRepo,
          branch: githubBaseBranch,
          filePath: targetPath,
        });
        const bytes = new Uint8Array(fetched.bytes);
        const probeSize = Math.min(bytes.byteLength, 1024 * 1024);
        const probeText = new TextDecoder().decode(bytes.subarray(0, probeSize));
        return probeText.includes("\r\n") ? "crlf" : "lf";
      } catch {
        return csvLineEnding;
      }
    },
    [csvLineEnding, githubBaseBranch, githubCsvFolder, githubRepo, githubToken],
  );

  const onOpenPullRequest = useCallback(async () => {
    setPrError("");
    if (!hasPublishableChanges) {
      setStatusMessage("No changes detected. Edit data before opening PR.");
      return;
    }
    if (!isGithubRepoValid) {
      setStatusMessage("Set a valid repository first (owner/name or URL).");
      return;
    }
    if (isExportingXlsx || isPublishingPr) {
      setStatusMessage("CSV generation is still in progress.");
      return;
    }

    setIsPublishingPr(true);
    try {
      const targetLineEnding = await detectRepoTargetLineEnding(csvExportFileName);
      const csvForPublish = await runCsvExport("publish", { saveToDisk: false, lineEnding: targetLineEnding });
      if (!csvForPublish) {
        setStatusMessage("CSV payload missing for publish.");
        return;
      }
      if (!githubToken) {
        setStatusMessage("Missing GitHub PAT token.");
        return;
      }

      const published = await publishCsvToGitHub({
        token: githubToken,
        repoFullName: githubRepo,
        baseBranch: githubBaseBranch,
        csvFolder: githubCsvFolder,
        fileName: csvForPublish.fileName,
        csvBytes: csvForPublish.bytes,
      });
      setPrUrl(published.prUrl);
      setPrError("");
      setStatusMessage(`PR created on ${githubRepo}: ${published.prUrl}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : "PR publish failed.";
      setPrError(message);
      setStatusMessage(message);
    } finally {
      setIsPublishingPr(false);
    }
  }, [
    csvExportFileName,
    detectRepoTargetLineEnding,
    githubBaseBranch,
    githubCsvFolder,
    githubRepo,
    githubToken,
    hasPublishableChanges,
    isExportingXlsx,
    isGithubRepoValid,
    isPublishingPr,
    runCsvExport,
  ]);

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
      `- Warnings (missing translations): ${validationFeedback.warningCount}`,
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

  const onRebuildTlk = useCallback(async () => {
    if (!rows.length || localeColumns.length === 0) {
      setStatusMessage("Import CSV first.");
      return;
    }

    const sourceNameForZip = csvExportFileName || mergedXlsxFile?.name || repoImportPath || "localization.csv";
    const zipName = makeRebuildZipName(sourceNameForZip);
    const zipRoot = makeSafeDownloadStem(safeFileNameFromPath(sourceNameForZip), "localization");
    const artifactStem = zipRoot;

    const generatedArtifacts: ArtifactRow[] = [];
    for (let i = 0; i < localeColumns.length; i += 1) {
      const col = localeColumns[i];
      const locale = String(col.locale || "").toUpperCase().trim();
      const languageId = TLK_LOCALE_TO_LANGUAGE_ID[locale] ?? 0;
      const binary = buildSingleTlkBinaryFromColumn(rows, col.field, languageId);
      const bytes = new Uint8Array(binary.byteLength);
      bytes.set(binary);
      const fileName = makeRebuildArtifactName(artifactStem, locale, col.variant === "dialogf");
      generatedArtifacts.push({
        file: fileName,
        checksum: quickChecksumHex(bytes),
        bytes,
        locale,
        encoding: tlkEncodingForLanguageId(languageId),
        isDialogf: col.variant === "dialogf",
        kind: "tlk",
      });
    }

    const downloadableArtifacts: ArtifactRow[] = [...generatedArtifacts];
    if (generatedArtifacts.length > 1) {
      const zipBytes = buildZipArchive(
        generatedArtifacts.map((artifact) => ({
          name: makeZipArtifactPath(zipRoot, artifactStem, artifact.locale, artifact.isDialogf),
          bytes: artifact.bytes,
        })),
      );
      downloadableArtifacts.unshift({
        file: zipName,
        checksum: quickChecksumHex(zipBytes),
        bytes: zipBytes,
        locale: "ALL",
        encoding: "bundle",
        isDialogf: false,
        kind: "zip",
      });
      setStatusMessage(`Rebuild completed. Bundle + ${generatedArtifacts.length} TLK files are ready to download.`);
    } else if (generatedArtifacts.length === 1) {
      setStatusMessage(`Rebuild completed. ${generatedArtifacts[0].file} is ready to download.`);
    } else {
      setStatusMessage("No TLK artifacts were generated.");
    }

    setArtifacts(downloadableArtifacts);
    setBuilt(true);
  }, [csvExportFileName, localeColumns, mergedXlsxFile?.name, repoImportPath, rows]);

  const onDownloadArtifact = useCallback(async (artifact: ArtifactRow) => {
    try {
      await saveBytesToDisk(artifact.bytes, artifact.file, "application/octet-stream");
      setStatusMessage(`Downloaded ${artifact.file}.`);
    } catch {
      setStatusMessage(`Could not download ${artifact.file}.`);
    }
  }, []);

  const onUndo = useCallback(() => {
    if (batchUndoStack.length > 0) {
      const [snapshot, ...rest] = batchUndoStack;
      setBatchUndoStack(rest);
      setBatchRedoStack((prev) => [cloneRows(rows), ...prev].slice(0, BATCH_HISTORY_LIMIT));
      applyRowsAndResetMeta(cloneRows(snapshot));
      setStatusMessage("Undid last bulk operation.");
      return;
    }
    if (!gridApi) return;
    gridApi.undoCellEditing();
  }, [applyRowsAndResetMeta, batchUndoStack, gridApi, rows]);

  const onRedo = useCallback(() => {
    if (batchRedoStack.length > 0) {
      const [snapshot, ...rest] = batchRedoStack;
      setBatchRedoStack(rest);
      setBatchUndoStack((prev) => [cloneRows(rows), ...prev].slice(0, BATCH_HISTORY_LIMIT));
      applyRowsAndResetMeta(cloneRows(snapshot));
      setStatusMessage("Redid last bulk operation.");
      return;
    }
    if (!gridApi) return;
    gridApi.redoCellEditing();
  }, [applyRowsAndResetMeta, batchRedoStack, gridApi, rows]);

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
        hasPublishableChanges,
      }),
    [activeStepIndex, exported, hasPublishableChanges, imported, scope, sourceLoaded, validated],
  );

  const canContinueExchange =
    sourceLoaded &&
    rows.length > 0 &&
    localeColumns.length > 0 &&
    hasPublishableChanges;

  const kpiItems = useMemo(
    () => [
      { value: String(emptyRows), label: "Empty rows" },
      { value: `${Math.round(coverage * 10) / 10}%`, label: "Coverage" },
      { value: currentPageSummary, label: "Dataset" },
    ],
    [coverage, currentPageSummary, emptyRows],
  );

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
      <header className="workflow-page__topbar workflow-page__topbar--scope-only">
        <ScopeSwitcher scope={scope} onScopeChange={onScopeChange} />
      </header>

      <StepStrip steps={stepTitles} activeStep={activeStepIndex} stepStates={stepStates} onStepSelect={onStepSelect} />

      <section className={`workflow-kpis${scope === "rebuild" ? " workflow-kpis--three" : ""}`}>
        {kpiItems.map((item) => (
          <article className="kpi" key={item.label}>
            <b>{item.value}</b>
            <small>{item.label}</small>
          </article>
        ))}
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
            renderCsvModelFilePicker({
              inputId: "xlsx-source",
              selectedFile: sourceXlsxFile,
              onFileChange: setSourceXlsxFile,
            })
          )}

          {importMode === "repo" && (
            renderRepoSourceFields()
          )}

          <div className="workflow-actions">
            <button type="button" onClick={onLoadSource} disabled={isLoadingSource}>
              {isLoadingSource ? "Loading..." : "Load Source"}
            </button>
            <button type="button" className="workflow-actions__primary" onClick={onParseAndGoEdit} disabled={!sourceLoaded || isLoadingSource}>
              Parse &amp; Validate
            </button>
          </div>
          {isLoadingSource ? (
            <div className="workflow-inline-progress" role="status" aria-live="polite">
              <span className="workflow-inline-progress__label">
                <span className="workflow-inline-progress__spinner" aria-hidden="true" />
                {loadSourceLabel || "Loading source..."}
              </span>
            </div>
          ) : null}
          {!isLoadingSource && loadSourceError ? (
            <div className="workflow-inline-error" role="alert" aria-live="assertive">
              {loadSourceError}
            </div>
          ) : null}
        </section>
      )}

      {scope === "exchange" && activeStepIndex === 1 && (
        <section className="workflow-screen">
          <header className="workflow-screen__header">
            <h2>2) Edit TLK Entries</h2>
            <p>Edit locale columns. Validate is optional before publish/rebuild.</p>
          </header>
          <div className="workflow-actions workflow-actions--editor-tools">
            <button
              type="button"
              className="workflow-actions__primary"
              onClick={applyQaValidation}
              disabled={!sourceLoaded || rows.length === 0 || localeColumns.length === 0}
              hidden
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
          {scope === "exchange" ? (
            <div className="workflow-screen__field workflow-screen__field--compact">
              <label htmlFor="csv-encoding">CSV encoding (for Export + Publish)</label>
              <select
                id="csv-encoding"
                value={csvEncoding}
                onChange={(event) => setCsvEncoding(event.target.value as CsvEncoding)}
              >
                {CSV_ENCODING_OPTIONS.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </div>
          ) : null}
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
              changedStrRefs={changeSummary.changedRowStrRefs}
              pageSize={pageSize}
              onPageSizeChange={setPageSize}
              onRowsChange={onGridRowsChange}
              onBatchRowsChange={onGridBatchRowsChange}
              onAddRow={onAddGridRow}
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
          <article className="card workflow-change-summary">
            <h3>Change Summary</h3>
            <div className="workflow-change-summary__grid">
              <div>
                <b>{changeSummary.changedRows}</b>
                <small>Changed rows</small>
              </div>
              <div>
                <b>{changeSummary.changedCells}</b>
                <small>Changed cells</small>
              </div>
              <div>
                <b>{changeSummary.addedRows}</b>
                <small>Added rows</small>
              </div>
              <div>
                <b>{changeSummary.removedRows}</b>
                <small>Removed rows</small>
              </div>
              <div>
                <b>{changeSummary.sourceChangedRows}</b>
                <small>Source EN updates</small>
              </div>
              <div>
                <b>{changeSummary.touchedLocales.length}</b>
                <small>Locales touched</small>
              </div>
            </div>
            {changeSummary.touchedLocales.length > 0 ? (
              <p className="workflow-screen__hint">{`Locales: ${changeSummary.touchedLocales.join(", ")}`}</p>
            ) : (
              <p className="workflow-screen__hint">No data changes detected yet.</p>
            )}
          </article>
          <div className="grid-2">
            <article className="card">
              <h3>GitHub PR</h3>
              <div className="workflow-screen__field">
                <label htmlFor="publish-repo-name">Repository (owner/name or URL)</label>
                <input
                  id="publish-repo-name"
                  type="text"
                  value={githubRepoInput}
                  onChange={(event) => onGithubRepoInputChange(event.target.value)}
                  placeholder={DEFAULT_GITHUB_REPO}
                />
              </div>
              <div className="workflow-screen__field">
                <label htmlFor="publish-branch">Base branch</label>
                <input
                  id="publish-branch"
                  type="text"
                  value={githubBaseBranchInput}
                  onChange={(event) => onGithubBaseBranchInputChange(event.target.value)}
                  placeholder={DEFAULT_GITHUB_BASE_BRANCH}
                />
              </div>
              <div className="workflow-screen__field">
                <label htmlFor="publish-csv-folder">CSV folder in repo</label>
                <input
                  id="publish-csv-folder"
                  type="text"
                  value={githubCsvFolderInput}
                  onChange={(event) => onGithubCsvFolderInputChange(event.target.value)}
                  placeholder={DEFAULT_GITHUB_CSV_FOLDER}
                />
              </div>
              <div className="workflow-screen__field">
                <label htmlFor="github-pat-token">GitHub PAT (test)</label>
                <input
                  id="github-pat-token"
                  type="password"
                  placeholder="github_pat_..."
                  value={githubRuntimeToken}
                  onChange={(event) => onGithubTokenChange(event.target.value)}
                />
              </div>
              <div className="workflow-actions">
                <button type="button" className="workflow-actions__secondary" onClick={() => onGithubTokenChange("")}>
                  Clear PAT
                </button>
              </div>
              <p className="workflow-screen__hint">
                Target repo:{" "}
                {isGithubRepoValid ? (
                  <a href={`https://github.com/${githubRepo}`} target="_blank" rel="noreferrer">
                    {githubRepo}
                  </a>
                ) : (
                  <span>Set valid repository first.</span>
                )}
              </p>
              <p className="workflow-screen__hint">
                CSV folder:{" "}
                {isGithubRepoValid ? (
                  <a href={githubCsvFolderUrl} target="_blank" rel="noreferrer">
                    {`${githubBaseBranch}/${githubCsvFolder}`}
                  </a>
                ) : (
                  <span>{`${githubBaseBranch}/${githubCsvFolder}`}</span>
                )}
              </p>
              <div className="workflow-actions">
                <button
                  type="button"
                  className="workflow-actions__primary"
                  onClick={onOpenPullRequest}
                  disabled={isExportingXlsx || isPublishingPr || !hasPublishableChanges || !githubToken || !isGithubRepoValid}
                >
                  {isPublishingPr ? "Opening PR..." : "Open PR"}
                </button>
              </div>
              {isPublishingPr ? (
                <div className="workflow-inline-progress" role="status" aria-live="polite">
                  <span className="workflow-inline-progress__label">
                    <span className="workflow-inline-progress__spinner" aria-hidden="true" />
                    {isExportingXlsx && exportRunMode === "publish"
                      ? `Preparing CSV for PR... ${exportProgressLabel || `${exportProgress}%`}`
                      : "Waiting for GitHub response..."}
                  </span>
                  {isExportingXlsx && exportRunMode === "publish" ? (
                    <div className="workflow-inline-progress__bar" aria-hidden="true">
                      <span style={{ width: `${Math.max(0, Math.min(100, exportProgress))}%` }} />
                    </div>
                  ) : null}
                </div>
              ) : null}
              {!isPublishingPr && prError ? (
                <div className="workflow-inline-error" role="alert" aria-live="assertive">
                  {`Open PR failed: ${prError}`}
                </div>
              ) : null}
              {!githubToken ? (
                <p className="workflow-screen__hint">Paste GitHub PAT to enable Open PR.</p>
              ) : null}
              {!isGithubRepoValid ? (
                <p className="workflow-screen__hint">Provide valid repository in format <code>owner/name</code> or GitHub URL.</p>
              ) : null}
              {!lastExport ? (
                <p className="workflow-screen__hint">CSV will be prepared automatically for PR (without local download).</p>
              ) : null}
              {!hasPublishableChanges ? (
                <p className="workflow-screen__hint">No data changes detected yet. Open PR is disabled.</p>
              ) : null}
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
                <button type="button" className="workflow-actions__primary" onClick={onExportXlsx} disabled={isExportingXlsx}>
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
            <h2>1) Load Approved CSV</h2>
          </header>
          <div className="mode-tabs">
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
          {importMode === "xlsx" && (
            renderCsvModelFilePicker({
              inputId: "merged-xlsx",
              selectedFile: mergedXlsxFile,
              onFileChange: setMergedXlsxFile,
            })
          )}
          {importMode === "repo" && renderRepoSourceFields()}
          <div className="workflow-actions">
            <button type="button" className="workflow-actions__primary" onClick={onLoadSource} disabled={isLoadingSource}>
              {isLoadingSource ? "Loading..." : "Load Source"}
            </button>
            <button
              type="button"
              className="workflow-actions__secondary"
              onClick={() => setStepIndex(1)}
              disabled={!imported || isLoadingSource}
            >
              Continue to Rebuild
            </button>
          </div>
          {isLoadingSource ? (
            <div className="workflow-inline-progress" role="status" aria-live="polite">
              <span className="workflow-inline-progress__label">
                <span className="workflow-inline-progress__spinner" aria-hidden="true" />
                {loadSourceLabel || "Importing source..."}
              </span>
            </div>
          ) : null}
          {!isLoadingSource && loadSourceError ? (
            <div className="workflow-inline-error" role="alert" aria-live="assertive">
              {loadSourceError}
            </div>
          ) : null}
        </section>
      )}

      {scope === "rebuild" && activeStepIndex === 1 && (
        <section className="workflow-screen">
          <header className="workflow-screen__header">
            <h2>2) Rebuild TLK/TLKF</h2>
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
                    <th>Encoding</th>
                    <th>Checksum</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {artifacts.map((artifact) => (
                    <tr key={artifact.file}>
                      <td>{artifact.file}</td>
                      <td>
                        <span className={statusBadgeClass("validated")}>{artifact.kind === "zip" ? "Bundle" : "Built"}</span>
                      </td>
                      <td>{artifact.encoding}</td>
                      <td>{artifact.checksum}</td>
                      <td>
                        <button type="button" onClick={() => void onDownloadArtifact(artifact)}>
                          {artifact.kind === "zip" ? "Download ZIP" : "Download"}
                        </button>
                      </td>
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
