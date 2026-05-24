import {
  type CellValueChangedEvent,
  type ColDef,
  type GridApi,
  type GridReadyEvent,
  type IRowNode,
} from "ag-grid-community";
import { AgGridReact } from "ag-grid-react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";

import type { LocaleColumn, TlkGridRow } from "../lib/types";

import "ag-grid-community/styles/ag-grid.css";
import "ag-grid-community/styles/ag-theme-quartz.css";

type LocalizationGridProps = {
  rows: TlkGridRow[];
  localeColumns: LocaleColumn[];
  changedStrRefs?: readonly number[];
  pageSize: number;
  onPageSizeChange: (nextPageSize: number) => void;
  onRowsChange: (rows: TlkGridRow[]) => void;
  onBatchRowsChange?: (rows: TlkGridRow[], previousRows?: TlkGridRow[]) => void;
  onAddRow?: () => number | null;
  onGridApiReady?: (api: GridApi<TlkGridRow>) => void;
  onUndo?: () => void;
  onRedo?: () => void;
};

type ColumnWidthPreset = "compact" | "balanced" | "wide";
type BulkCopyOption = {
  field: string;
  label: string;
};

const pageOptions = [10, 25, 50, 100, 200, 500];
const MIN_GRID_HEIGHT = 220;
const DEFAULT_GRID_HEIGHT = 420;
const DEFAULT_COLUMN_WIDTH_PRESET: ColumnWidthPreset = "balanced";
const COLUMN_WIDTH_PRESETS: Record<
  ColumnWidthPreset,
  { strRef: number; locale: number }
> = {
  compact: { strRef: 95, locale: 170 },
  balanced: { strRef: 110, locale: 200 },
  wide: { strRef: 125, locale: 260 },
};

const LocalizationGrid = ({
  rows,
  localeColumns,
  changedStrRefs = [],
  pageSize,
  onPageSizeChange,
  onRowsChange,
  onBatchRowsChange,
  onAddRow,
  onGridApiReady,
  onUndo,
  onRedo,
}: LocalizationGridProps) => {
  const [gridApi, setGridApi] = useState<GridApi<TlkGridRow> | null>(null);
  const [pagingInfo, setPagingInfo] = useState("Page 1 / 1");
  const [visibleRowsInfo, setVisibleRowsInfo] = useState("Visible 0 / 0");
  const [searchText, setSearchText] = useState("");
  const [notEmptyOnly, setNotEmptyOnly] = useState(false);
  const [changedOnly, setChangedOnly] = useState(false);
  const [missingOnly, setMissingOnly] = useState(false);
  const [targetEqualsSourceOnly, setTargetEqualsSourceOnly] = useState(false);
  const [sourceNotEmptyTargetEmptyOnly, setSourceNotEmptyTargetEmptyOnly] = useState(false);
  const [showShortcuts, setShowShortcuts] = useState(false);
  const [gridHeight, setGridHeight] = useState(DEFAULT_GRID_HEIGHT);
  const [gridHeightDraft, setGridHeightDraft] = useState(String(DEFAULT_GRID_HEIGHT));
  const [gridHeightError, setGridHeightError] = useState("");
  const [columnWidthPreset, setColumnWidthPreset] = useState<ColumnWidthPreset>(DEFAULT_COLUMN_WIDTH_PRESET);
  const [columnWidthError, setColumnWidthError] = useState("");
  const [copyFeedback, setCopyFeedback] = useState("");
  const [bulkCopySafeMode, setBulkCopySafeMode] = useState(true);
  const [bulkCopyFilledSourceOnly, setBulkCopyFilledSourceOnly] = useState(true);
  const [findText, setFindText] = useState("");
  const [replaceText, setReplaceText] = useState("");
  const [findReplaceScope, setFindReplaceScope] = useState<"target" | "all">("target");
  const [bulkCopyFrom, setBulkCopyFrom] = useState("");
  const [bulkCopyTo, setBulkCopyTo] = useState("");
  const gridResizeRef = useRef<HTMLDivElement | null>(null);
  const copyFeedbackTimeoutRef = useRef<number | null>(null);
  const changedStrRefSet = useMemo(() => new Set(changedStrRefs.map((value) => Number(value))), [changedStrRefs]);
  const bulkCopyOptions = useMemo<BulkCopyOption[]>(() => {
    return localeColumns.map((col) => ({ field: col.field, label: col.title }));
  }, [localeColumns]);
  const validBulkCopyFrom = useMemo(() => {
    if (bulkCopyOptions.some((option) => option.field === bulkCopyFrom)) {
      return bulkCopyFrom;
    }
    return bulkCopyOptions[0]?.field || "";
  }, [bulkCopyFrom, bulkCopyOptions]);
  const validBulkCopyTo = useMemo(() => {
    const fromField = validBulkCopyFrom;
    const hasCurrent = bulkCopyOptions.some((option) => option.field === bulkCopyTo);
    if (hasCurrent && bulkCopyTo !== fromField) {
      return bulkCopyTo;
    }
    const fallback = bulkCopyOptions.find((option) => option.field !== fromField);
    return fallback?.field || "";
  }, [bulkCopyOptions, bulkCopyTo, validBulkCopyFrom]);

  const columnDefs = useMemo<ColDef<TlkGridRow>[]>(() => {
    const localeDefs: ColDef<TlkGridRow>[] = localeColumns.map((col) => ({
      field: col.field,
      headerName: col.title,
      flex: 1.4,
      minWidth: 180,
      editable: true,
    }));

    return [
      { field: "strRef", headerName: "StrRef", width: 110, pinned: "left" },
      ...localeDefs,
    ];
  }, [localeColumns]);

  const defaultColDef = useMemo<ColDef<TlkGridRow>>(
    () => ({
      sortable: true,
      filter: true,
      resizable: true,
      editable: false,
      wrapText: false,
    }),
    [],
  );

  const updatePaging = useCallback(() => {
    if (!gridApi) return;
    const current = gridApi.paginationGetCurrentPage() + 1;
    const total = gridApi.paginationGetTotalPages();
    setPagingInfo(`Page ${current} / ${total || 1}`);
    let visibleRows = 0;
    gridApi.forEachNodeAfterFilter(() => {
      visibleRows += 1;
    });
    setVisibleRowsInfo(`Visible ${visibleRows} / ${rows.length}`);
  }, [gridApi, rows.length]);

  const handleGridReady = useCallback(
    (event: GridReadyEvent<TlkGridRow>) => {
      const api = event.api;
      setGridApi(api);
      updatePaging();
      api.setGridOption("paginationPageSize", pageSize);
      onGridApiReady?.(api);
    },
    [onGridApiReady, pageSize, updatePaging],
  );

  useEffect(() => {
    if (!gridApi) return;
    gridApi.setGridOption("paginationPageSize", pageSize);
    updatePaging();
  }, [gridApi, pageSize, updatePaging]);

  useEffect(() => {
    if (!gridApi) return;
    gridApi.setGridOption("quickFilterText", searchText);
    gridApi.onFilterChanged();
    updatePaging();
  }, [gridApi, searchText, updatePaging]);

  useEffect(() => {
    if (!gridApi) return;
    gridApi.onFilterChanged();
    updatePaging();
  }, [
    changedOnly,
    gridApi,
    missingOnly,
    notEmptyOnly,
    sourceNotEmptyTargetEmptyOnly,
    targetEqualsSourceOnly,
    updatePaging,
    validBulkCopyFrom,
    validBulkCopyTo,
  ]);

  useEffect(() => {
    const host = gridResizeRef.current;
    if (!host || typeof ResizeObserver === "undefined") return;

    const observer = new ResizeObserver(() => {
      // Use border-box height so typed value does not drift down by border thickness.
      const nextHeight = Math.max(MIN_GRID_HEIGHT, Math.round(host.offsetHeight));
      if (!Number.isFinite(nextHeight) || nextHeight <= 0) return;
      setGridHeight((prev) => (Math.abs(prev - nextHeight) > 1 ? nextHeight : prev));
    });

    observer.observe(host);
    return () => observer.disconnect();
  }, []);

  useEffect(() => {
    setGridHeightDraft(String(gridHeight));
  }, [gridHeight]);

  const resolveWidthForColumn = useCallback((colId: string, preset: ColumnWidthPreset): number => {
    const config = COLUMN_WIDTH_PRESETS[preset];
    if (colId === "strRef") return config.strRef;
    if (colId.startsWith("loc_")) return config.locale;
    return config.locale;
  }, []);

  const applyColumnWidthPreset = useCallback(
    (preset: ColumnWidthPreset) => {
      if (!gridApi) return;
      const columns = gridApi.getColumns();
      if (!columns || columns.length === 0) return;

      const state = columns.map((column) => ({
        colId: column.getColId(),
        width: resolveWidthForColumn(column.getColId(), preset),
      }));

      const ok = gridApi.applyColumnState({ state });
      if (!ok) {
        setColumnWidthError("Could not apply width to all columns.");
        return;
      }
      setColumnWidthError("");
    },
    [gridApi, resolveWidthForColumn],
  );

  const resetColumnWidths = useCallback(() => {
    if (!gridApi) return;
    gridApi.resetColumnState();
    setColumnWidthPreset(DEFAULT_COLUMN_WIDTH_PRESET);
    setColumnWidthError("");
  }, [gridApi]);

  useEffect(() => {
    if (!gridApi) return;
    applyColumnWidthPreset(columnWidthPreset);
  }, [applyColumnWidthPreset, columnWidthPreset, columnDefs, gridApi]);

  useEffect(() => {
    return () => {
      if (copyFeedbackTimeoutRef.current != null) {
        window.clearTimeout(copyFeedbackTimeoutRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (validBulkCopyFrom !== bulkCopyFrom) {
      setBulkCopyFrom(validBulkCopyFrom);
    }
    if (validBulkCopyTo !== bulkCopyTo) {
      setBulkCopyTo(validBulkCopyTo);
    }
  }, [bulkCopyFrom, bulkCopyTo, validBulkCopyFrom, validBulkCopyTo]);

  const setTransientCopyFeedback = useCallback((message: string) => {
    setCopyFeedback(message);
    if (copyFeedbackTimeoutRef.current != null) {
      window.clearTimeout(copyFeedbackTimeoutRef.current);
    }
    copyFeedbackTimeoutRef.current = window.setTimeout(() => {
      setCopyFeedback("");
      copyFeedbackTimeoutRef.current = null;
    }, 1800);
  }, []);

  const copyTextToClipboard = useCallback(async (text: string): Promise<boolean> => {
    if (typeof navigator !== "undefined" && navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }

    if (typeof document === "undefined") return false;
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.setAttribute("readonly", "");
    textArea.style.position = "absolute";
    textArea.style.left = "-9999px";
    document.body.appendChild(textArea);
    textArea.select();
    textArea.setSelectionRange(0, textArea.value.length);
    const copied = document.execCommand("copy");
    textArea.remove();
    return copied;
  }, []);

  const copyFocusedCellValue = useCallback(async () => {
    if (!gridApi) {
      setTransientCopyFeedback("Grid not ready.");
      return;
    }

    const focusedCell = gridApi.getFocusedCell();
    if (!focusedCell) {
      setTransientCopyFeedback("Select a cell first.");
      return;
    }

    const rowNode = gridApi.getDisplayedRowAtIndex(focusedCell.rowIndex);
    const row = rowNode?.data;
    if (!row) {
      setTransientCopyFeedback("No row selected.");
      return;
    }

    const colId = focusedCell.column.getColId();
    const value = String((row as Record<string, unknown>)[colId] ?? "");

    try {
      const copied = await copyTextToClipboard(value);
      setTransientCopyFeedback(copied ? "Copied cell value." : "Could not copy cell value.");
    } catch {
      setTransientCopyFeedback("Clipboard blocked by browser.");
    }
  }, [copyTextToClipboard, gridApi, setTransientCopyFeedback]);

  const applyBulkCopyColumns = useCallback(() => {
    if (!validBulkCopyFrom || !validBulkCopyTo || validBulkCopyFrom === validBulkCopyTo) {
      setTransientCopyFeedback("Select different source and target columns.");
      return;
    }

    const fromLabel = bulkCopyOptions.find((option) => option.field === validBulkCopyFrom)?.label || validBulkCopyFrom;
    const toLabel = bulkCopyOptions.find((option) => option.field === validBulkCopyTo)?.label || validBulkCopyTo;
    let changedCount = 0;

    const nextRows = rows.map((row) => {
      const sourceValue = String((row as Record<string, unknown>)[validBulkCopyFrom] ?? "");
      const targetValue = String((row as Record<string, unknown>)[validBulkCopyTo] ?? "");
      const sourceTrimmed = sourceValue.trim();
      const targetTrimmed = targetValue.trim();
      const sourceAllowed = bulkCopyFilledSourceOnly ? sourceTrimmed.length > 0 : true;
      const targetAllowed = bulkCopySafeMode ? targetTrimmed.length === 0 : true;
      if (!sourceAllowed || !targetAllowed) {
        return row;
      }
      if (sourceValue !== targetValue) {
        changedCount += 1;
      }
      return {
        ...row,
        [validBulkCopyTo]: sourceValue,
      };
    });

    if (changedCount === 0) {
      setTransientCopyFeedback(`No changes (${fromLabel} -> ${toLabel}).`);
      return;
    }

    if (onBatchRowsChange) {
      onBatchRowsChange(nextRows);
    } else {
      onRowsChange(nextRows);
    }
    setTransientCopyFeedback(`Copied ${changedCount} row(s): ${fromLabel} -> ${toLabel}.`);
  }, [
    bulkCopyFilledSourceOnly,
    bulkCopyOptions,
    bulkCopySafeMode,
    onBatchRowsChange,
    onRowsChange,
    rows,
    setTransientCopyFeedback,
    validBulkCopyFrom,
    validBulkCopyTo,
  ]);

  const applyFindReplace = useCallback(() => {
    const find = findText;
    if (!find) {
      setTransientCopyFeedback("Enter text to find.");
      return;
    }
    const columns =
      findReplaceScope === "all"
        ? localeColumns.map((col) => col.field)
        : validBulkCopyTo
          ? [validBulkCopyTo]
          : [];
    if (columns.length === 0) {
      setTransientCopyFeedback("Select target column first.");
      return;
    }

    let changedCells = 0;
    const nextRows = rows.map((row) => {
      let rowChanged = false;
      const nextRow: TlkGridRow = { ...row };
      for (let i = 0; i < columns.length; i += 1) {
        const field = columns[i];
        const currentValue = String((row as Record<string, unknown>)[field] ?? "");
        if (!currentValue.includes(find)) continue;
        const replacedValue = currentValue.split(find).join(replaceText);
        if (replacedValue === currentValue) continue;
        nextRow[field] = replacedValue;
        changedCells += 1;
        rowChanged = true;
      }
      return rowChanged ? nextRow : row;
    });

    if (changedCells === 0) {
      setTransientCopyFeedback("No matching cells.");
      return;
    }
    if (onBatchRowsChange) {
      onBatchRowsChange(nextRows);
    } else {
      onRowsChange(nextRows);
    }
    setTransientCopyFeedback(`Find & replace updated ${changedCells} cell(s).`);
  }, [
    findReplaceScope,
    findText,
    localeColumns,
    onBatchRowsChange,
    onRowsChange,
    replaceText,
    rows,
    setTransientCopyFeedback,
    validBulkCopyTo,
  ]);

  useEffect(() => {
    const isTypingTarget = (target: EventTarget | null) => {
      const element = target as HTMLElement | null;
      if (!element) return false;
      const tag = String(element.tagName || "").toLowerCase();
      if (tag === "input" || tag === "textarea" || tag === "select") return true;
      return element.isContentEditable;
    };

    const onKeyDown = (event: KeyboardEvent) => {
      const ctrlOrMeta = event.ctrlKey || event.metaKey;
      const lower = event.key.toLowerCase();

      if (event.key === "Escape" && showShortcuts) {
        event.preventDefault();
        setShowShortcuts(false);
        return;
      }

      if (ctrlOrMeta && lower === "z" && !event.shiftKey && onUndo) {
        event.preventDefault();
        onUndo();
        return;
      }

      if (ctrlOrMeta && ((lower === "y") || (lower === "z" && event.shiftKey)) && onRedo) {
        event.preventDefault();
        onRedo();
        return;
      }

      if (isTypingTarget(event.target)) {
        return;
      }

      if (ctrlOrMeta && lower === "c") {
        const hasFocusedGridCell = Boolean(gridApi?.getFocusedCell());
        if (hasFocusedGridCell) {
          event.preventDefault();
          void copyFocusedCellValue();
          return;
        }
      }

      if (event.key === "?") {
        event.preventDefault();
        setShowShortcuts((prev) => !prev);
        return;
      }

      if (event.altKey && lower === "1") {
        event.preventDefault();
        setNotEmptyOnly((prev) => !prev);
        return;
      }
      if (event.altKey && lower === "2") {
        event.preventDefault();
        setMissingOnly((prev) => !prev);
        return;
      }
      if (event.altKey && lower === "3") {
        event.preventDefault();
        setChangedOnly((prev) => !prev);
        return;
      }
      if (event.altKey && lower === "c") {
        event.preventDefault();
        void copyFocusedCellValue();
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [changedOnly, copyFocusedCellValue, gridApi, onRedo, onUndo, showShortcuts]);

  const isExternalFilterPresent = useCallback(() => {
    return notEmptyOnly || changedOnly || missingOnly || targetEqualsSourceOnly || sourceNotEmptyTargetEmptyOnly;
  }, [changedOnly, missingOnly, notEmptyOnly, sourceNotEmptyTargetEmptyOnly, targetEqualsSourceOnly]);

  const doesExternalFilterPass = useCallback(
    (node: IRowNode<TlkGridRow>) => {
      const row = node.data;
      if (!row) return false;

      const hasValue = localeColumns.some((col) => String(row[col.field] || "").trim().length > 0);
      const hasMissing = localeColumns.some((col) => String(row[col.field] || "").trim().length === 0);
      const sourceValue = String((row as Record<string, unknown>)[validBulkCopyFrom] ?? "").trim();
      const targetValue = String((row as Record<string, unknown>)[validBulkCopyTo] ?? "").trim();

      if (notEmptyOnly && !hasValue) return false;
      if (changedOnly && !changedStrRefSet.has(Number(row.strRef))) return false;
      if (missingOnly && !hasMissing) return false;
      if (targetEqualsSourceOnly) {
        if (!sourceValue || !targetValue || sourceValue !== targetValue) return false;
      }
      if (sourceNotEmptyTargetEmptyOnly) {
        if (!sourceValue || targetValue) return false;
      }

      return true;
    },
    [
      changedOnly,
      changedStrRefSet,
      localeColumns,
      missingOnly,
      notEmptyOnly,
      sourceNotEmptyTargetEmptyOnly,
      targetEqualsSourceOnly,
      validBulkCopyFrom,
      validBulkCopyTo,
    ],
  );

  const resetQuickFilters = useCallback(() => {
    setSearchText("");
    setNotEmptyOnly(false);
    setChangedOnly(false);
    setMissingOnly(false);
    setTargetEqualsSourceOnly(false);
    setSourceNotEmptyTargetEmptyOnly(false);
  }, []);

  const resetMergedToolbarState = useCallback(() => {
    resetQuickFilters();
    resetColumnWidths();
  }, [resetColumnWidths, resetQuickFilters]);

  const handleAddRow = useCallback(() => {
    const createdStrRef = onAddRow?.();
    if (!gridApi || createdStrRef == null) {
      return;
    }

    window.requestAnimationFrame(() => {
      gridApi.paginationGoToLastPage();
      updatePaging();

      const displayedCount = gridApi.getDisplayedRowCount();
      if (displayedCount <= 0) return;

      const targetIndex = displayedCount - 1;
      gridApi.ensureIndexVisible(targetIndex, "bottom");
      const node = gridApi.getDisplayedRowAtIndex(targetIndex);
      node?.setSelected(true);
    });
  }, [gridApi, onAddRow, updatePaging]);

  const applyGridHeightDraft = useCallback(() => {
    const normalized = gridHeightDraft.trim();
    const parsed = Number(normalized);
    if (!normalized || !Number.isFinite(parsed)) {
      setGridHeightError("Enter a valid numeric height.");
      return;
    }
    const nextHeight = Math.round(parsed);
    if (nextHeight < MIN_GRID_HEIGHT) {
      setGridHeightError(`Grid height must be at least ${MIN_GRID_HEIGHT}px.`);
      return;
    }
    setGridHeightError("");
    setGridHeight(nextHeight);
  }, [gridHeightDraft]);

  const handleCellValueChanged = useCallback(
    (event: CellValueChangedEvent<TlkGridRow>) => {
      if (!event.data) return;
      const changedField = event.colDef.field;
      const changedStrRef = Number(event.data.strRef);
      const nextRows: TlkGridRow[] = [];
      event.api.forEachNode((node) => {
        if (!node.data) return;
        const row = { ...node.data };
        nextRows.push(row);
      });

      let previousRowsSnapshot: TlkGridRow[] | undefined;
      if (changedField && event.oldValue !== event.newValue) {
        previousRowsSnapshot = nextRows.map((row) => ({ ...row }));
        const targetRow =
          previousRowsSnapshot.find((row) => Number(row.strRef) === changedStrRef) ??
          (typeof event.rowIndex === "number" &&
          event.rowIndex >= 0 &&
          event.rowIndex < previousRowsSnapshot.length
            ? previousRowsSnapshot[event.rowIndex]
            : undefined);
        if (targetRow) {
          (targetRow as Record<string, unknown>)[changedField] = event.oldValue as unknown;
        }
      }

      if (onBatchRowsChange) {
        onBatchRowsChange(nextRows, previousRowsSnapshot);
        return;
      }
      onRowsChange(nextRows);
    },
    [onBatchRowsChange, onRowsChange],
  );

  return (
    <section className="workflow-grid">
      <header className="workflow-grid__toolbar">
        <div className="workflow-grid__toolbar-main">
          <div className="workflow-grid__toolbar-row workflow-grid__toolbar-row--top">
            <label className="workflow-grid__search">
              <span>Search</span>
              <input
                type="text"
                value={searchText}
                placeholder="StrRef, locale..."
                onChange={(event) => setSearchText(event.target.value)}
              />
            </label>
            <div className="workflow-grid__settings">
              <label className="workflow-grid__page-size">
                <span>Rows per page</span>
                <select value={pageSize} onChange={(event) => onPageSizeChange(Number(event.target.value))}>
                  {pageOptions.map((option) => (
                    <option key={option} value={option}>
                      {option}
                    </option>
                  ))}
                </select>
              </label>
              <label className="workflow-grid__column-width">
                <span>Column width</span>
                <div className="workflow-grid__column-width-controls">
                  <select
                    value={columnWidthPreset}
                    onChange={(event) => setColumnWidthPreset(event.target.value as ColumnWidthPreset)}
                  >
                    <option value="compact">Compact</option>
                    <option value="balanced">Balanced</option>
                    <option value="wide">Wide</option>
                  </select>
                </div>
              </label>
              <label className="workflow-grid__height">
                <span>Grid height</span>
                <div className="workflow-grid__height-input-wrap">
                  <input
                    type="text"
                    inputMode="numeric"
                    value={gridHeightDraft}
                    onChange={(event) => {
                      setGridHeightDraft(event.target.value);
                      if (gridHeightError) setGridHeightError("");
                    }}
                    onKeyDown={(event) => {
                      if (event.key === "Enter") {
                        event.preventDefault();
                        applyGridHeightDraft();
                      }
                    }}
                  />
                  <button type="button" className="workflow-grid__toolbar-action workflow-grid__toolbar-action--compact" onClick={applyGridHeightDraft}>
                    Apply
                  </button>
                  <small>px</small>
                </div>
              </label>
            </div>
          </div>
          <div className="workflow-grid__toolbar-row workflow-grid__toolbar-row--bottom">
            <div className="workflow-grid__filters">
              <button
                type="button"
                aria-pressed={notEmptyOnly}
                className={`workflow-grid__chip ${notEmptyOnly ? "workflow-grid__chip--active" : ""}`}
                onClick={() => setNotEmptyOnly((prev) => !prev)}
              >
                Not Empty
              </button>
              <button
                type="button"
                aria-pressed={changedOnly}
                className={`workflow-grid__chip ${changedOnly ? "workflow-grid__chip--active" : ""}`}
                onClick={() => setChangedOnly((prev) => !prev)}
                disabled={changedStrRefs.length === 0}
              >
                Review Changes Only
              </button>
              <button
                type="button"
                aria-pressed={missingOnly}
                className={`workflow-grid__chip ${missingOnly ? "workflow-grid__chip--active" : ""}`}
                onClick={() => setMissingOnly((prev) => !prev)}
              >
                Missing Translations
              </button>
              <button
                type="button"
                aria-pressed={targetEqualsSourceOnly}
                className={`workflow-grid__chip ${targetEqualsSourceOnly ? "workflow-grid__chip--active" : ""}`}
                onClick={() => setTargetEqualsSourceOnly((prev) => !prev)}
                disabled={!validBulkCopyFrom || !validBulkCopyTo || validBulkCopyFrom === validBulkCopyTo}
              >
                Target == Source
              </button>
              <button
                type="button"
                aria-pressed={sourceNotEmptyTargetEmptyOnly}
                className={`workflow-grid__chip ${sourceNotEmptyTargetEmptyOnly ? "workflow-grid__chip--active" : ""}`}
                onClick={() => setSourceNotEmptyTargetEmptyOnly((prev) => !prev)}
                disabled={!validBulkCopyFrom || !validBulkCopyTo || validBulkCopyFrom === validBulkCopyTo}
              >
                Source filled + Target empty
              </button>
              <button type="button" className="workflow-grid__chip workflow-grid__chip--reset" onClick={resetMergedToolbarState}>
                Reset
              </button>
            </div>
          </div>
          <div className="workflow-grid__toolbar-row workflow-grid__toolbar-row--bulk-copy">
            <div className="workflow-grid__bulk-copy">
              <div className="workflow-grid__bulk-copy-main">
                <span>Copy Column</span>
                <select
                  value={validBulkCopyFrom}
                  onChange={(event) => setBulkCopyFrom(event.target.value)}
                  aria-label="Bulk copy source column"
                >
                  {bulkCopyOptions.map((option) => (
                    <option key={`copy-from-${option.field}`} value={option.field}>
                      {option.label}
                    </option>
                  ))}
                </select>
                <small>to</small>
                <select
                  value={validBulkCopyTo}
                  onChange={(event) => setBulkCopyTo(event.target.value)}
                  aria-label="Bulk copy target column"
                >
                  {bulkCopyOptions.map((option) => (
                    <option key={`copy-to-${option.field}`} value={option.field}>
                      {option.label}
                    </option>
                  ))}
                </select>
                <button
                  type="button"
                  className="workflow-grid__toolbar-action workflow-grid__toolbar-action--compact"
                  onClick={applyBulkCopyColumns}
                  disabled={
                    rows.length === 0 ||
                    !validBulkCopyFrom ||
                    !validBulkCopyTo ||
                    validBulkCopyFrom === validBulkCopyTo
                  }
                >
                  Copy All Rows
                </button>
              </div>
              <div className="workflow-grid__bulk-copy-toggles">
                <label className="workflow-grid__inline-check">
                  <input
                    type="checkbox"
                    checked={bulkCopySafeMode}
                    onChange={(event) => setBulkCopySafeMode(event.target.checked)}
                  />
                  <span>Safe copy (target empty only)</span>
                </label>
                <label className="workflow-grid__inline-check">
                  <input
                    type="checkbox"
                    checked={bulkCopyFilledSourceOnly}
                    onChange={(event) => setBulkCopyFilledSourceOnly(event.target.checked)}
                  />
                  <span>Copy filled source only</span>
                </label>
              </div>
            </div>
          </div>
          <div className="workflow-grid__toolbar-row workflow-grid__toolbar-row--find-replace">
            <div className="workflow-grid__find-replace">
              <span>Find &amp; Replace</span>
              <input
                type="text"
                value={findText}
                onChange={(event) => setFindText(event.target.value)}
                placeholder="Find text..."
                aria-label="Find text"
              />
              <small>→</small>
              <input
                type="text"
                value={replaceText}
                onChange={(event) => setReplaceText(event.target.value)}
                placeholder="Replace with..."
                aria-label="Replace text"
              />
              <select
                value={findReplaceScope}
                onChange={(event) => setFindReplaceScope(event.target.value as "target" | "all")}
                aria-label="Find replace scope"
              >
                <option value="target">Target column</option>
                <option value="all">All locale columns</option>
              </select>
              <button
                type="button"
                className="workflow-grid__toolbar-action workflow-grid__toolbar-action--compact"
                onClick={applyFindReplace}
                disabled={rows.length === 0 || findText.length === 0}
              >
                Apply Replace
              </button>
            </div>
          </div>
          <div className="workflow-grid__toolbar-row workflow-grid__toolbar-row--info">
            <div className="workflow-grid__history">
              <button
                type="button"
                className="workflow-grid__toolbar-action"
                onClick={handleAddRow}
                disabled={!onAddRow}
              >
                Add Row
              </button>
              <button
                type="button"
                className="workflow-grid__toolbar-action"
                onClick={() => {
                  void copyFocusedCellValue();
                }}
                disabled={!gridApi}
              >
                Copy Cell
              </button>
              <button
                type="button"
                className="workflow-grid__toolbar-action"
                onClick={() => onUndo?.()}
                disabled={!onUndo}
              >
                Undo
              </button>
              <button
                type="button"
                className="workflow-grid__toolbar-action"
                onClick={() => onRedo?.()}
                disabled={!onRedo}
              >
                Redo
              </button>
            </div>
            <div className="workflow-grid__toolbar-status">
              {gridHeightError ? <small className="workflow-grid__toolbar-error">{gridHeightError}</small> : null}
              {columnWidthError ? <small className="workflow-grid__toolbar-error">{columnWidthError}</small> : null}
              {copyFeedback ? <small className="workflow-grid__toolbar-copy">{copyFeedback}</small> : null}
              <button
                type="button"
                className="workflow-grid__toolbar-action workflow-grid__toolbar-action--compact"
                onClick={() => setShowShortcuts((prev) => !prev)}
                aria-expanded={showShortcuts}
                aria-controls="workflow-grid-shortcuts"
                title="Show shortcuts"
              >
                ?
              </button>
              <span className="workflow-grid__meta">{`${visibleRowsInfo} | ${pagingInfo}`}</span>
            </div>
          </div>
        </div>
      </header>
      {showShortcuts ? (
        <aside id="workflow-grid-shortcuts" className="workflow-grid__shortcuts" aria-live="polite">
          <h4>Keyboard Shortcuts</h4>
          <ul>
            <li><code>Ctrl/Cmd + Z</code> Undo</li>
            <li><code>Ctrl/Cmd + Y</code> or <code>Ctrl/Cmd + Shift + Z</code> Redo</li>
            <li><code>?</code> Toggle this help</li>
            <li><code>Alt + 1</code> Not Empty</li>
            <li><code>Alt + 2</code> Missing Translations</li>
            <li><code>Alt + 3</code> Review Changes Only</li>
            <li><code>Target == Source</code> Uses selected copy source/target columns</li>
            <li><code>Source filled + Target empty</code> Uses selected copy source/target columns</li>
            <li><code>Alt + C</code> Copy focused cell</li>
            <li><code>Esc</code> Close this help</li>
          </ul>
        </aside>
      ) : null}
      <div className="workflow-grid__resizable" ref={gridResizeRef} style={{ height: `${gridHeight}px` }}>
        <div className="ag-theme-quartz workflow-grid__canvas">
          <AgGridReact
            rowData={rows}
            columnDefs={columnDefs}
            defaultColDef={defaultColDef}
            pagination
            paginationPageSize={pageSize}
            paginationPageSizeSelector={false}
            undoRedoCellEditing
            undoRedoCellEditingLimit={100}
            rowSelection="multiple"
            onGridReady={handleGridReady}
            onPaginationChanged={updatePaging}
            onFilterChanged={updatePaging}
            onCellValueChanged={handleCellValueChanged}
            isExternalFilterPresent={isExternalFilterPresent}
            doesExternalFilterPass={doesExternalFilterPass}
          />
        </div>
      </div>
    </section>
  );
};

export default LocalizationGrid;
