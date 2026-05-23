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
  onAddRow?: () => number | null;
  onGridApiReady?: (api: GridApi<TlkGridRow>) => void;
  onUndo?: () => void;
  onRedo?: () => void;
};

type ColumnWidthPreset = "compact" | "balanced" | "wide";

const pageOptions = [10, 25, 50, 100, 200, 500];
const MIN_GRID_HEIGHT = 220;
const DEFAULT_GRID_HEIGHT = 420;
const DEFAULT_COLUMN_WIDTH_PRESET: ColumnWidthPreset = "balanced";
const COLUMN_WIDTH_PRESETS: Record<
  ColumnWidthPreset,
  { strRef: number; sourceEn: number; locale: number; status: number }
> = {
  compact: { strRef: 95, sourceEn: 220, locale: 170, status: 120 },
  balanced: { strRef: 110, sourceEn: 260, locale: 200, status: 130 },
  wide: { strRef: 125, sourceEn: 320, locale: 260, status: 150 },
};

const LocalizationGrid = ({
  rows,
  localeColumns,
  changedStrRefs = [],
  pageSize,
  onPageSizeChange,
  onRowsChange,
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
  const [needsQaOnly, setNeedsQaOnly] = useState(false);
  const [errorsOnly, setErrorsOnly] = useState(false);
  const [showShortcuts, setShowShortcuts] = useState(false);
  const [gridHeight, setGridHeight] = useState(DEFAULT_GRID_HEIGHT);
  const [gridHeightDraft, setGridHeightDraft] = useState(String(DEFAULT_GRID_HEIGHT));
  const [gridHeightError, setGridHeightError] = useState("");
  const [columnWidthPreset, setColumnWidthPreset] = useState<ColumnWidthPreset>(DEFAULT_COLUMN_WIDTH_PRESET);
  const [columnWidthError, setColumnWidthError] = useState("");
  const gridResizeRef = useRef<HTMLDivElement | null>(null);
  const changedStrRefSet = useMemo(() => new Set(changedStrRefs.map((value) => Number(value))), [changedStrRefs]);

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
      { field: "sourceEn", headerName: "Source EN", minWidth: 260, flex: 1.8, pinned: "left" },
      ...localeDefs,
      {
        field: "status",
        headerName: "Status",
        width: 130,
        editable: false,
        valueFormatter: (params) => {
          const value = String(params.value || "draft");
          return value[0].toUpperCase() + value.slice(1);
        },
      },
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
  }, [changedOnly, errorsOnly, gridApi, missingOnly, needsQaOnly, notEmptyOnly, updatePaging]);

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
    if (colId === "sourceEn") return config.sourceEn;
    if (colId === "status") return config.status;
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
        setNeedsQaOnly((prev) => !prev);
        return;
      }
      if (event.altKey && lower === "4") {
        event.preventDefault();
        setErrorsOnly((prev) => !prev);
        return;
      }
      if (event.altKey && lower === "5") {
        event.preventDefault();
        setChangedOnly((prev) => !prev);
      }
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [onRedo, onUndo, showShortcuts]);

  const isExternalFilterPresent = useCallback(() => {
    return notEmptyOnly || changedOnly || missingOnly || needsQaOnly || errorsOnly;
  }, [changedOnly, errorsOnly, missingOnly, needsQaOnly, notEmptyOnly]);

  const doesExternalFilterPass = useCallback(
    (node: IRowNode<TlkGridRow>) => {
      const row = node.data;
      if (!row) return false;

      const hasValue = localeColumns.some((col) => String(row[col.field] || "").trim().length > 0);
      const hasMissing = localeColumns.some((col) => String(row[col.field] || "").trim().length === 0);
      const status = String(row.status || "").toLowerCase();

      if (notEmptyOnly && !hasValue) return false;
      if (changedOnly && !changedStrRefSet.has(Number(row.strRef))) return false;
      if (missingOnly && !hasMissing) return false;
      if (needsQaOnly && status !== "needs qa") return false;
      if (errorsOnly && status !== "error") return false;

      return true;
    },
    [changedOnly, changedStrRefSet, errorsOnly, localeColumns, missingOnly, needsQaOnly, notEmptyOnly],
  );

  const resetQuickFilters = useCallback(() => {
    setSearchText("");
    setNotEmptyOnly(false);
    setChangedOnly(false);
    setMissingOnly(false);
    setNeedsQaOnly(false);
    setErrorsOnly(false);
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
      const nextRows: TlkGridRow[] = [];
      event.api.forEachNode((node) => {
        if (!node.data) return;
        const row = { ...node.data };
        nextRows.push(row);
      });
      onRowsChange(nextRows);
    },
    [onRowsChange],
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
                placeholder="StrRef, source, locale..."
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
                aria-pressed={needsQaOnly}
                className={`workflow-grid__chip ${needsQaOnly ? "workflow-grid__chip--active" : ""}`}
                onClick={() => setNeedsQaOnly((prev) => !prev)}
              >
                Needs QA
              </button>
              <button
                type="button"
                aria-pressed={errorsOnly}
                className={`workflow-grid__chip ${errorsOnly ? "workflow-grid__chip--active" : ""}`}
                onClick={() => setErrorsOnly((prev) => !prev)}
              >
                Errors
              </button>
              <button type="button" className="workflow-grid__chip workflow-grid__chip--reset" onClick={resetMergedToolbarState}>
                Reset
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
            <li><code>Alt + 3</code> Needs QA</li>
            <li><code>Alt + 4</code> Errors</li>
            <li><code>Alt + 5</code> Review Changes Only</li>
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
