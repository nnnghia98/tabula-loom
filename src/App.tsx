import React, { useState, useRef, useMemo, useEffect, useCallback } from 'react';
import type {
  WorkbookProject,
  SavedWorkbookSummary,
  SheetData,
  FieldConfig,
  CellValue,
  WorkspaceMode,
  DbStatus,
  ActiveCellInfo,
  LookupOption,
} from './types';
import { EMPTY_WORKBOOK } from './sampleData';
import { loadDraft, saveDraft, clearDraft } from './draftStorage';
import {
  parseXlsxFile,
  buildSheetFromMatrix,
  autoDetectCrossSheetLookups,
  validateSheet,
  getExcelColumnLetter,
} from './xlsxUtils';
import { arrayBufferToBase64 } from './utils/base64';
import { normalizeCellValue, buildDefaultRowCells } from './utils/cellValues';
import { createBlankField, createBlankSheet } from './utils/sheetFactory';
import { useColumnWidths } from './hooks/useColumnWidths';
import { TopNavbar } from './components/TopNavbar';
import { DraftBanner } from './components/DraftBanner';
import { HiddenSheetsBanner } from './components/HiddenSheetsBanner';
import { FormulaBar } from './components/FormulaBar';
import { SheetGrid } from './components/SheetGrid';
import { FieldInspector } from './components/FieldInspector';
import { SchemaView } from './components/SchemaView';
import { SheetTabsBar } from './components/SheetTabsBar';
import { DbFilesDialog } from './components/DbFilesDialog';
import { ExportDialog } from './components/ExportDialog';
import { NewSheetDialog } from './components/NewSheetDialog';
import { AddRowDialog } from './components/AddRowDialog';

export function App() {
  const [project, setProject] = useState<WorkbookProject>(EMPTY_WORKBOOK);
  const [rawUploadedBase64, setRawUploadedBase64] = useState<string | null>(null);
  const [activeSheetId, setActiveSheetId] = useState<string>(EMPTY_WORKBOOK.sheets[0].id);
  const [workspaceMode, setWorkspaceMode] = useState<WorkspaceMode>('grid');
  const [selectedFieldId, setSelectedFieldId] = useState<string>(
    EMPTY_WORKBOOK.sheets[0].fields[0].id
  );
  const [selectedCell, setSelectedCell] = useState<{ rowId: string; fieldId: string } | null>({
    rowId: EMPTY_WORKBOOK.sheets[0].rows[0].id,
    fieldId: EMPTY_WORKBOOK.sheets[0].fields[0].id,
  });
  const [editingCell, setEditingCell] = useState<{ rowId: string; fieldId: string } | null>(null);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [showHiddenColumns, setShowHiddenColumns] = useState<boolean>(false);
  const [showHiddenSheets, setShowHiddenSheets] = useState<boolean>(false);
  const [dismissedHiddenNotice, setDismissedHiddenNotice] = useState<boolean>(false);

  // Accessibility & Simplicity State for Excel users
  const [uiScale, setUiScale] = useState<number>(1.0); // 1.0 (Normal), 1.15 (Large), 1.3 (Extra Large)
  const [isInspectorOpen, setIsInspectorOpen] = useState<boolean>(false); // Closed by default so spreadsheet has 100% width

  // PostgreSQL & Storage Backend State
  const [dbStatus, setDbStatus] = useState<DbStatus>({
    checked: false,
    connected: false,
    error: null,
    storageMode: 'volume-or-disk',
  });
  const [savedWorkbooks, setSavedWorkbooks] = useState<SavedWorkbookSummary[]>([]);
  const [isSavingToDb, setIsSavingToDb] = useState<boolean>(false);
  const [isLoadingFromDb, setIsLoadingFromDb] = useState<boolean>(false);
  const [saveToastMessage, setSaveToastMessage] = useState<string | null>(null);

  // Export dialog options
  const [exportUseMachineKeys, setExportUseMachineKeys] = useState<boolean>(false);
  const [exportIncludeHidden, setExportIncludeHidden] = useState<boolean>(false);
  const [exportIncludeHiddenSheets, setExportIncludeHiddenSheets] = useState<boolean>(false);
  const [exportIncludeSchemaTab, setExportIncludeSchemaTab] = useState<boolean>(true);

  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const exportDialogRef = useRef<HTMLDialogElement | null>(null);
  const newSheetDialogRef = useRef<HTMLDialogElement | null>(null);
  const dbDrawerDialogRef = useRef<HTMLDialogElement | null>(null);
  const addRowDialogRef = useRef<HTMLDialogElement | null>(null);
  const [newSheetName, setNewSheetName] = useState<string>('');
  const [newRowDraft, setNewRowDraft] = useState<Record<string, CellValue>>({});
  const [newRowErrors, setNewRowErrors] = useState<Record<string, string>>({});

  const fetchBackendStatusAndWorkbooks = useCallback(async () => {
    try {
      const [healthRes, listRes] = await Promise.all([
        fetch('/api/health'),
        fetch('/api/workbooks'),
      ]);
      if (healthRes.ok) {
        const health = await healthRes.json();
        setDbStatus({
          checked: true,
          connected: Boolean(health.dbConnected),
          error: health.dbError || null,
          storageMode: health.storageMode || 'volume-or-disk',
        });
      }
      if (listRes.ok) {
        const listData = await listRes.json();
        setSavedWorkbooks(listData.workbooks || []);
      }
    } catch {
      setDbStatus((prev) => ({
        ...prev,
        checked: true,
        connected: false,
        error: 'Backend server is not connected.',
      }));
    }
  }, []);

  useEffect(() => {
    fetchBackendStatusAndWorkbooks();
  }, [fetchBackendStatusAndWorkbooks]);

  const showToast = (msg: string) => {
    setSaveToastMessage(msg);
    setTimeout(() => {
      setSaveToastMessage((prev) => (prev === msg ? null : prev));
    }, 3600);
  };

  // Local draft (IndexedDB): restore unsaved work after a tab close/reload.
  const [draftState, setDraftState] = useState<'idle' | 'saved'>('idle');
  const [restoredDraftAt, setRestoredDraftAt] = useState<number | null>(null);
  const draftReadyRef = useRef(false);
  const skipNextDraftWriteRef = useRef(false);

  useEffect(() => {
    let cancelled = false;
    loadDraft().then((draft) => {
      if (cancelled) return;
      if (draft && draft.project?.sheets?.length > 0) {
        skipNextDraftWriteRef.current = true;
        setProject(draft.project);
        setRawUploadedBase64(draft.rawUploadedBase64);
        const sheet =
          draft.project.sheets.find((s) => s.id === draft.activeSheetId) ||
          draft.project.sheets[0];
        setActiveSheetId(sheet.id);
        setSelectedFieldId(sheet.fields[0]?.id || '');
        setSelectedCell(
          sheet.rows[0] && sheet.fields[0]
            ? { rowId: sheet.rows[0].id, fieldId: sheet.fields[0].id }
            : null
        );
        setRestoredDraftAt(draft.savedAt);
        setDraftState('saved');
      }
      draftReadyRef.current = true;
    });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!draftReadyRef.current) return;
    if (skipNextDraftWriteRef.current) {
      skipNextDraftWriteRef.current = false;
      return;
    }
    if (project === EMPTY_WORKBOOK) return;
    const timer = setTimeout(async () => {
      const ok = await saveDraft({
        project,
        rawUploadedBase64,
        activeSheetId,
        savedAt: Date.now(),
      });
      setDraftState(ok ? 'saved' : 'idle');
    }, 800);
    return () => clearTimeout(timer);
  }, [project, rawUploadedBase64, activeSheetId]);

  const handleDiscardDraft = async () => {
    await clearDraft();
    skipNextDraftWriteRef.current = true;
    setProject(EMPTY_WORKBOOK);
    setRawUploadedBase64(null);
    setActiveSheetId(EMPTY_WORKBOOK.sheets[0].id);
    setSelectedFieldId(EMPTY_WORKBOOK.sheets[0].fields[0].id);
    setSelectedCell({
      rowId: EMPTY_WORKBOOK.sheets[0].rows[0].id,
      fieldId: EMPTY_WORKBOOK.sheets[0].fields[0].id,
    });
    setRestoredDraftAt(null);
    setDraftState('idle');
  };

  const handleSaveToPostgres = async () => {
    setIsSavingToDb(true);
    try {
      const res = await fetch('/api/workbooks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          id: project.id,
          fileName: project.fileName,
          sheets: project.sheets,
          rawFileBase64: rawUploadedBase64 || undefined,
        }),
      });
      const data = await res.json();
      if (!res.ok) {
        dbDrawerDialogRef.current?.showModal();
        showToast(data.error || 'Could not save workbook');
        return;
      }
      skipNextDraftWriteRef.current = true;
      setProject((prev) => ({
        ...prev,
        id: data.id,
        lastModified: data.lastModified,
        hasOriginalBinary: prev.hasOriginalBinary || Boolean(rawUploadedBase64),
      }));
      setRawUploadedBase64(null);
      setRestoredDraftAt(null);
      setDraftState('idle');
      await clearDraft();
      await fetchBackendStatusAndWorkbooks();
      showToast(`Saved "${project.fileName}"`);
    } catch {
      dbDrawerDialogRef.current?.showModal();
      showToast('Could not connect to database');
    } finally {
      setIsSavingToDb(false);
    }
  };

  const handleLoadWorkbookFromDb = async (workbookId: string) => {
    setIsLoadingFromDb(true);
    try {
      const res = await fetch(`/api/workbooks/${workbookId}`);
      if (!res.ok) {
        showToast('Could not open file');
        return;
      }
      const loaded: WorkbookProject = await res.json();
      if (loaded.sheets.length > 0) {
        skipNextDraftWriteRef.current = true;
        setProject(loaded);
        setRawUploadedBase64(null);
        setRestoredDraftAt(null);
        void clearDraft();
        setDraftState('idle');
        setDismissedHiddenNotice(false);
        const firstVisible = loaded.sheets.find((s) => !s.hiddenState) || loaded.sheets[0];
        setActiveSheetId(firstVisible.id);
        setSelectedFieldId(firstVisible.fields[0]?.id || '');
        setSelectedCell(
          firstVisible.rows[0] && firstVisible.fields[0]
            ? { rowId: firstVisible.rows[0].id, fieldId: firstVisible.fields[0].id }
            : null
        );
        dbDrawerDialogRef.current?.close();
        showToast(`Opened "${loaded.fileName}"`);
      }
    } finally {
      setIsLoadingFromDb(false);
    }
  };

  const handleDeleteWorkbookFromDb = async (workbookId: string, fileName: string) => {
    const res = await fetch(`/api/workbooks/${workbookId}`, { method: 'DELETE' });
    if (res.ok) {
      await fetchBackendStatusAndWorkbooks();
      if (project.id === workbookId) {
        setProject((prev) => ({ ...prev, id: undefined }));
      }
      showToast(`Deleted "${fileName}"`);
    }
  };

  const hiddenSheets = useMemo(
    () => project.sheets.filter((s) => Boolean(s.hiddenState)),
    [project.sheets]
  );

  const displayedSheets = useMemo(() => {
    if (showHiddenSheets) return project.sheets;
    const visibleOnly = project.sheets.filter((s) => !s.hiddenState);
    return visibleOnly.length > 0 ? visibleOnly : project.sheets;
  }, [project.sheets, showHiddenSheets]);

  const activeSheet = useMemo(
    () =>
      displayedSheets.find((s) => s.id === activeSheetId) ||
      project.sheets.find((s) => s.id === activeSheetId) ||
      displayedSheets[0] ||
      project.sheets[0],
    [displayedSheets, project.sheets, activeSheetId]
  );

  const selectedField = useMemo(
    () =>
      activeSheet.fields.find((f) => f.id === selectedFieldId) ||
      activeSheet.fields[0] ||
      null,
    [activeSheet, selectedFieldId]
  );

  // Compute live validation errors for all sheets
  const sheetErrorsMap = useMemo(() => {
    const map: Record<string, ReturnType<typeof validateSheet>> = {};
    project.sheets.forEach((s) => {
      map[s.id] = validateSheet(s, project.sheets);
    });
    return map;
  }, [project.sheets]);

  const activeSheetErrors = sheetErrorsMap[activeSheet.id] || {};
  const totalErrorCount = Object.values(sheetErrorsMap).reduce(
    (acc, errs) => acc + Object.keys(errs).length,
    0
  );

  const visibleFields = useMemo(() => {
    if (showHiddenColumns) return activeSheet.fields;
    return activeSheet.fields.filter((f) => f.visible);
  }, [activeSheet.fields, showHiddenColumns]);

  const { getColWidth, startColumnResize } = useColumnWidths(activeSheet);

  const filteredRows = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return activeSheet.rows;
    return activeSheet.rows.filter((row) =>
      activeSheet.fields.some((f) => {
        const val = row.cells[f.id];
        return val !== null && val !== undefined && String(val).toLowerCase().includes(q);
      })
    );
  }, [activeSheet, searchQuery]);

  // Handlers
  const handleSelectSheet = (sheetId: string) => {
    setActiveSheetId(sheetId);
    const target = project.sheets.find((s) => s.id === sheetId);
    if (target && target.fields.length > 0) {
      setSelectedFieldId(target.fields[0].id);
      if (target.rows.length > 0) {
        setSelectedCell({ rowId: target.rows[0].id, fieldId: target.fields[0].id });
      } else {
        setSelectedCell(null);
      }
    }
    setEditingCell(null);
  };

  const loadWorkbookBuffer = (buffer: ArrayBuffer, fileName: string) => {
    const parsed = parseXlsxFile(buffer, fileName);
    if (parsed.sheets.length > 0) {
      setProject(parsed);
      setRawUploadedBase64(arrayBufferToBase64(buffer));
      setDismissedHiddenNotice(false);
      const firstVisible = parsed.sheets.find((s) => !s.hiddenState) || parsed.sheets[0];
      setActiveSheetId(firstVisible.id);
      setSelectedFieldId(firstVisible.fields[0]?.id || '');
      setSelectedCell(
        firstVisible.rows[0] && firstVisible.fields[0]
          ? { rowId: firstVisible.rows[0].id, fieldId: firstVisible.fields[0].id }
          : null
      );
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    loadWorkbookBuffer(buffer, file.name);
    e.target.value = '';
  };

  const handleHeaderRowChange = (newHeaderRowIdx: number) => {
    const sourceMatrix: CellValue[][] =
      activeSheet.rawMatrix ||
      [
        activeSheet.fields.map((f) => f.label),
        ...activeSheet.rows.map((r) => activeSheet.fields.map((f) => r.cells[f.id] ?? '')),
      ];

    const rebuiltSheet = buildSheetFromMatrix(
      activeSheet.id,
      activeSheet.name,
      sourceMatrix,
      newHeaderRowIdx,
      activeSheet.hiddenState || 0
    );

    setProject((prev) => {
      const nextSheets = prev.sheets.map((s) => (s.id === activeSheet.id ? rebuiltSheet : s));
      const linkedSheets = autoDetectCrossSheetLookups(nextSheets);
      return {
        ...prev,
        lastModified: new Date().toISOString(),
        sheets: linkedSheets,
      };
    });

    if (rebuiltSheet.fields[0]) {
      setSelectedFieldId(rebuiltSheet.fields[0].id);
      if (rebuiltSheet.rows[0]) {
        setSelectedCell({
          rowId: rebuiltSheet.rows[0].id,
          fieldId: rebuiltSheet.fields[0].id,
        });
      }
    }
  };

  const handlePromoteDataRowToHeader = (rowId: string) => {
    const rowIdx = activeSheet.rows.findIndex((r) => r.id === rowId);
    if (rowIdx < 0) return;

    const targetRow = activeSheet.rows[rowIdx];
    if (activeSheet.rawMatrix && targetRow.excelRowNumber !== undefined) {
      handleHeaderRowChange(targetRow.excelRowNumber - 1);
      return;
    }

    const newHeaderArray: CellValue[] = activeSheet.fields.map(
      (f) => targetRow.cells[f.id] ?? f.label
    );
    const remainingDataMatrix: CellValue[][] = activeSheet.rows
      .slice(rowIdx + 1)
      .map((r) => activeSheet.fields.map((f) => r.cells[f.id] ?? ''));

    const matrix: CellValue[][] = [newHeaderArray, ...remainingDataMatrix];
    const rebuilt = buildSheetFromMatrix(
      activeSheet.id,
      activeSheet.name,
      matrix,
      0,
      activeSheet.hiddenState || 0
    );

    setProject((prev) => ({
      ...prev,
      lastModified: new Date().toISOString(),
      sheets: prev.sheets.map((s) => (s.id === activeSheet.id ? rebuilt : s)),
    }));
  };

  const updateActiveSheet = (updater: (sheet: SheetData) => SheetData) => {
    setProject((prev) => ({
      ...prev,
      lastModified: new Date().toISOString(),
      sheets: prev.sheets.map((s) => (s.id === activeSheet.id ? updater(s) : s)),
    }));
  };

  const updateFieldConfig = (fieldId: string, patch: Partial<FieldConfig>) => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      fields: sheet.fields.map((f) => (f.id === fieldId ? { ...f, ...patch } : f)),
    }));
  };

  const updateCellValue = (rowId: string, fieldId: string, rawValue: CellValue) => {
    const field = activeSheet.fields.find((f) => f.id === fieldId);
    const normalized = normalizeCellValue(field, rawValue);

    updateActiveSheet((sheet) => ({
      ...sheet,
      rows: sheet.rows.map((r) =>
        r.id === rowId ? { ...r, cells: { ...r.cells, [fieldId]: normalized } } : r
      ),
    }));
  };

  const buildRowCells = () => buildDefaultRowCells(activeSheet.fields);

  const handleOpenAddRowModal = () => {
    setNewRowDraft(buildRowCells());
    setNewRowErrors({});
    addRowDialogRef.current?.showModal();
  };

  const insertNewRowWithCells = (cellsPayload: Record<string, CellValue>) => {
    const newRowId = `row_${Date.now().toString(36)}`;
    const normalizedCells: Record<string, CellValue> = {};

    activeSheet.fields.forEach((f) => {
      const rawValue = cellsPayload[f.id] ?? '';
      if (f.type === 'number' && typeof rawValue === 'string' && rawValue.trim() !== '') {
        const cleaned = rawValue.replace(/[$,%\s]/g, '');
        const parsedNum = Number(cleaned);
        normalizedCells[f.id] = Number.isNaN(parsedNum) ? rawValue : parsedNum;
      } else if (f.type === 'boolean') {
        normalizedCells[f.id] =
          rawValue === true || String(rawValue).toLowerCase() === 'true';
      } else {
        normalizedCells[f.id] = rawValue;
      }
    });

    updateActiveSheet((sheet) => ({
      ...sheet,
      rows: [
        ...sheet.rows,
        {
          id: newRowId,
          excelRowNumber:
            (sheet.rows[sheet.rows.length - 1]?.excelRowNumber || sheet.rows.length) + 1,
          cells: normalizedCells,
        },
      ],
    }));

    if (activeSheet.fields[0]) {
      setSelectedCell({ rowId: newRowId, fieldId: activeSheet.fields[0].id });
    }
  };

  const handleAddBlankRow = () => {
    insertNewRowWithCells(buildRowCells());
    addRowDialogRef.current?.close();
  };

  const handleSubmitAddRowForm = (e: React.FormEvent, keepOpen = false) => {
    e.preventDefault();
    const tempRowId = '__new_row__';
    const tempSheet: SheetData = {
      ...activeSheet,
      rows: [
        ...activeSheet.rows,
        {
          id: tempRowId,
          cells: newRowDraft,
        },
      ],
    };
    const sheetValidation = validateSheet(tempSheet, project.sheets);
    const rowFieldErrors: Record<string, string> = {};
    activeSheet.fields.forEach((f) => {
      const err = sheetValidation[`${tempRowId}:${f.id}`];
      if (err) {
        rowFieldErrors[f.id] = err.message;
      }
    });

    if (Object.keys(rowFieldErrors).length > 0) {
      setNewRowErrors(rowFieldErrors);
      return;
    }

    insertNewRowWithCells(newRowDraft);
    if (keepOpen) {
      setNewRowDraft(buildRowCells());
      setNewRowErrors({});
      showToast(`Added row to "${activeSheet.name}"`);
    } else {
      addRowDialogRef.current?.close();
    }
  };

  const handleDeleteRow = (rowId: string) => {
    updateActiveSheet((sheet) => ({
      ...sheet,
      rows: sheet.rows.filter((r) => r.id !== rowId),
    }));
  };

  const handleAddField = () => {
    const newField = createBlankField(activeSheet.fields.length, activeSheet.headerRowIndex === -1);

    updateActiveSheet((sheet) => ({
      ...sheet,
      fields: [...sheet.fields, newField],
      rows: sheet.rows.map((r) => ({
        ...r,
        cells: { ...r.cells, [newField.id]: '' },
      })),
    }));
    setSelectedFieldId(newField.id);
  };

  const handleDeleteField = (fieldId: string) => {
    if (activeSheet.fields.length <= 1) return;
    const remaining = activeSheet.fields.filter((f) => f.id !== fieldId);
    updateActiveSheet((sheet) => ({
      ...sheet,
      fields: remaining,
    }));
    if (selectedFieldId === fieldId && remaining[0]) {
      setSelectedFieldId(remaining[0].id);
    }
  };

  const handleMoveField = (fieldId: string, direction: -1 | 1) => {
    const idx = activeSheet.fields.findIndex((f) => f.id === fieldId);
    if (idx < 0) return;
    const targetIdx = idx + direction;
    if (targetIdx < 0 || targetIdx >= activeSheet.fields.length) return;
    const nextFields = [...activeSheet.fields];
    const [moved] = nextFields.splice(idx, 1);
    nextFields.splice(targetIdx, 0, moved);
    updateActiveSheet((sheet) => ({
      ...sheet,
      fields: nextFields,
    }));
  };

  const handleCreateSheet = (e: React.FormEvent) => {
    e.preventDefault();
    const trimmed = newSheetName.trim() || `Sheet ${project.sheets.length + 1}`;
    const newSheet = createBlankSheet(trimmed);

    setProject((prev) => ({
      ...prev,
      sheets: [...prev.sheets, newSheet],
    }));
    setActiveSheetId(newSheet.id);
    setSelectedFieldId(newSheet.fields[0].id);
    setNewSheetName('');
    newSheetDialogRef.current?.close();
  };

  // Lookup helper for rendering foreign key options & resolved display labels
  const getLookupOptions = (field: FieldConfig) => {
    if (field.type !== 'lookup' || !field.lookup) return [];
    const targetSheet = project.sheets.find((s) => s.id === field.lookup?.targetSheetId);
    if (!targetSheet) return [];
    return targetSheet.rows.map((r) => {
      const keyVal = String(r.cells[field.lookup!.targetFieldId] ?? '');
      const displayVal = field.lookup?.displayFieldId
        ? String(r.cells[field.lookup.displayFieldId] ?? '')
        : '';
      return { keyVal, displayVal, sheetName: targetSheet.name };
    });
  };

  // Active cell coordinate label & value for formula bar
  const activeCellInfo = useMemo(() => {
    if (!selectedCell) return null;
    const rowIdx = activeSheet.rows.findIndex((r) => r.id === selectedCell.rowId);
    const colIdx = visibleFields.findIndex((f) => f.id === selectedCell.fieldId);
    const field = activeSheet.fields.find((f) => f.id === selectedCell.fieldId);
    const row = activeSheet.rows[rowIdx];
    if (!field || !row) return null;

    const colLetter = field.columnLetter || getExcelColumnLetter(Math.max(0, colIdx));
    const displayRowNum = row.excelRowNumber ?? rowIdx + 1;
    const coord = `${colLetter}${displayRowNum}`;
    const rawValue = row.cells[field.id];
    const err = activeSheetErrors[`${row.id}:${field.id}`];

    return { coord, field, row, rawValue, err, rowIdx, displayRowNum };
  }, [selectedCell, activeSheet, visibleFields, activeSheetErrors]);

  const isHeaderlessSheet = activeSheet.headerRowIndex === -1;

  return (
    <div
      className="app-shell"
      style={{ '--ui-scale': uiScale } as React.CSSProperties}
    >
      {/* 1. Simple, Readable Top Bar */}
      <TopNavbar
        activeSheet={activeSheet}
        activeSheetErrors={activeSheetErrors}
        dbDrawerDialogRef={dbDrawerDialogRef}
        draftState={draftState}
        exportDialogRef={exportDialogRef}
        fetchBackendStatusAndWorkbooks={fetchBackendStatusAndWorkbooks}
        fileInputRef={fileInputRef}
        handleFileUpload={handleFileUpload}
        handleSaveToPostgres={handleSaveToPostgres}
        isSavingToDb={isSavingToDb}
        project={project}
        saveToastMessage={saveToastMessage}
        savedWorkbooks={savedWorkbooks}
        setProject={setProject}
        setSelectedCell={setSelectedCell}
        setSelectedFieldId={setSelectedFieldId}
        setUiScale={setUiScale}
        setWorkspaceMode={setWorkspaceMode}
        totalErrorCount={totalErrorCount}
        uiScale={uiScale}
        workspaceMode={workspaceMode}
      />

      {/* Restored Local Draft Banner */}
      <DraftBanner
        handleDiscardDraft={handleDiscardDraft}
        restoredDraftAt={restoredDraftAt}
        setRestoredDraftAt={setRestoredDraftAt}
      />

      {/* Hidden Sheets Detection Notice Banner */}
      <HiddenSheetsBanner
        dismissedHiddenNotice={dismissedHiddenNotice}
        hiddenSheets={hiddenSheets}
        setDismissedHiddenNotice={setDismissedHiddenNotice}
        setShowHiddenSheets={setShowHiddenSheets}
        showHiddenSheets={showHiddenSheets}
      />

      {/* Main Workspace Body */}
      <div className="workspace-body">
        {workspaceMode === 'grid' ? (
          <>
            <main className="primary-stage">
              <FormulaBar
                activeCellInfo={activeCellInfo}
                handleAddField={handleAddField}
                handleOpenAddRowModal={handleOpenAddRowModal}
                isInspectorOpen={isInspectorOpen}
                searchQuery={searchQuery}
                setIsInspectorOpen={setIsInspectorOpen}
                setSearchQuery={setSearchQuery}
                updateCellValue={updateCellValue}
                updateFieldConfig={updateFieldConfig}
              />
              <SheetGrid
                activeSheet={activeSheet}
                activeSheetErrors={activeSheetErrors}
                editingCell={editingCell}
                filteredRows={filteredRows}
                getColWidth={getColWidth}
                getLookupOptions={getLookupOptions}
                handleDeleteRow={handleDeleteRow}
                isHeaderlessSheet={isHeaderlessSheet}
                selectedCell={selectedCell}
                selectedFieldId={selectedFieldId}
                setEditingCell={setEditingCell}
                setIsInspectorOpen={setIsInspectorOpen}
                setSelectedCell={setSelectedCell}
                setSelectedFieldId={setSelectedFieldId}
                startColumnResize={startColumnResize}
                updateCellValue={updateCellValue}
                visibleFields={visibleFields}
              />
            </main>

            {isInspectorOpen && selectedField && (
              <FieldInspector
                activeSheet={activeSheet}
                handleDeleteField={handleDeleteField}
                handleMoveField={handleMoveField}
                project={project}
                selectedField={selectedField}
                setIsInspectorOpen={setIsInspectorOpen}
                updateFieldConfig={updateFieldConfig}
              />
            )}
          </>
        ) : (
          <SchemaView
            activeSheet={activeSheet}
            handleAddField={handleAddField}
            handleDeleteField={handleDeleteField}
            handleHeaderRowChange={handleHeaderRowChange}
            project={project}
            selectedFieldId={selectedFieldId}
            setSelectedFieldId={setSelectedFieldId}
            setShowHiddenColumns={setShowHiddenColumns}
            setWorkspaceMode={setWorkspaceMode}
            showHiddenColumns={showHiddenColumns}
            updateFieldConfig={updateFieldConfig}
          />
        )}
      </div>

      {/* Classic Excel Bottom Bar */}
      <SheetTabsBar
        activeSheet={activeSheet}
        dismissedHiddenNotice={dismissedHiddenNotice}
        displayedSheets={displayedSheets}
        filteredRows={filteredRows}
        handleSelectSheet={handleSelectSheet}
        hiddenSheets={hiddenSheets}
        newSheetDialogRef={newSheetDialogRef}
        setShowHiddenSheets={setShowHiddenSheets}
        sheetErrorsMap={sheetErrorsMap}
        showHiddenSheets={showHiddenSheets}
        visibleFields={visibleFields}
      />

      <DbFilesDialog
        dbDrawerDialogRef={dbDrawerDialogRef}
        dbStatus={dbStatus}
        handleDeleteWorkbookFromDb={handleDeleteWorkbookFromDb}
        handleLoadWorkbookFromDb={handleLoadWorkbookFromDb}
        isLoadingFromDb={isLoadingFromDb}
        project={project}
        savedWorkbooks={savedWorkbooks}
      />

      <ExportDialog
        activeSheet={activeSheet}
        exportDialogRef={exportDialogRef}
        exportIncludeHidden={exportIncludeHidden}
        exportIncludeHiddenSheets={exportIncludeHiddenSheets}
        exportIncludeSchemaTab={exportIncludeSchemaTab}
        exportUseMachineKeys={exportUseMachineKeys}
        hiddenSheets={hiddenSheets}
        project={project}
        setExportIncludeHidden={setExportIncludeHidden}
        setExportIncludeHiddenSheets={setExportIncludeHiddenSheets}
        setExportIncludeSchemaTab={setExportIncludeSchemaTab}
      />

      <NewSheetDialog
        handleCreateSheet={handleCreateSheet}
        newSheetDialogRef={newSheetDialogRef}
        newSheetName={newSheetName}
        setNewSheetName={setNewSheetName}
      />

      <AddRowDialog
        activeSheet={activeSheet}
        addRowDialogRef={addRowDialogRef}
        getLookupOptions={getLookupOptions}
        handleAddBlankRow={handleAddBlankRow}
        handleSubmitAddRowForm={handleSubmitAddRowForm}
        newRowDraft={newRowDraft}
        newRowErrors={newRowErrors}
        setNewRowDraft={setNewRowDraft}
        setNewRowErrors={setNewRowErrors}
      />
    </div>
  );
}
