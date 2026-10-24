import React, { useState, useRef, useMemo, useEffect, useCallback } from 'react';
import {
  WorkbookProject,
  SavedWorkbookSummary,
  SheetData,
  FieldConfig,
  FieldDataType,
  NumberFormatType,
  CellValue,
} from './types';
import { EMPTY_WORKBOOK } from './sampleData';
import { loadDraft, saveDraft, clearDraft } from './draftStorage';
import {
  parseXlsxFile,
  buildSheetFromMatrix,
  autoDetectCrossSheetLookups,
  validateSheet,
  formatCellValue,
  exportWorkbookToXlsx,
  exportSchemaAsJson,
  toMachineKey,
  getExcelColumnLetter,
} from './xlsxUtils';
import {
  IconUpload,
  IconDownload,
  IconPlus,
  IconTrash,
  IconEye,
  IconSearch,
  IconSliders,
  IconTable,
  IconLink,
  IconWarning,
  IconCheck,
  IconDatabase,
  FieldTypeGlyph,
} from './Icons';

type WorkspaceMode = 'grid' | 'schema';

const COL_WIDTHS_STORAGE_KEY = 'tabula-loom:column-widths';

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

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
  const [dbStatus, setDbStatus] = useState<{
    checked: boolean;
    connected: boolean;
    error: string | null;
    storageMode: string;
  }>({
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

  // Persisted per-column width overrides (localStorage), keyed by `${sheetId}:${fieldId}`
  const [colWidths, setColWidths] = useState<Record<string, number>>(() => {
    try {
      return JSON.parse(localStorage.getItem(COL_WIDTHS_STORAGE_KEY) || '{}');
    } catch {
      return {};
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem(COL_WIDTHS_STORAGE_KEY, JSON.stringify(colWidths));
    } catch {
      /* storage unavailable */
    }
  }, [colWidths]);

  const isAutoIdField = (field: FieldConfig, colIdx: number): boolean => {
    if (/^(stt|no\.?|#|id)$/i.test(field.label.trim())) return true;
    if (colIdx !== 0 || activeSheet.rows.length < 2) return false;
    const nonEmpty = activeSheet.rows
      .map((r) => r.cells[field.id])
      .filter((v) => v !== null && v !== undefined && String(v).trim() !== '');
    return (
      nonEmpty.length >= 2 &&
      nonEmpty.every((v, idx) => Number(String(v).trim()) === idx + 1)
    );
  };

  const getColWidth = (field: FieldConfig, colIdx = 0): number => {
    const manualWidth = colWidths[`${activeSheet.id}:${field.id}`];
    if (manualWidth !== undefined) return manualWidth;
    if (isAutoIdField(field, colIdx)) return 64;
    return field.width;
  };

  const startColumnResize = (e: React.MouseEvent, field: FieldConfig, colIdx = 0) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = getColWidth(field, colIdx);
    const key = `${activeSheet.id}:${field.id}`;
    const onMove = (ev: MouseEvent) => {
      const next = Math.max(48, Math.round(startWidth + ev.clientX - startX));
      setColWidths((prev) => ({ ...prev, [key]: next }));
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
      document.body.style.userSelect = '';
      document.body.style.cursor = '';
    };
    document.body.style.userSelect = 'none';
    document.body.style.cursor = 'col-resize';
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  };

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
    let normalized: CellValue = rawValue;
    if (field?.type === 'number' && typeof rawValue === 'string' && rawValue.trim() !== '') {
      const cleaned = rawValue.replace(/[$,%\s]/g, '');
      const parsedNum = Number(cleaned);
      normalized = Number.isNaN(parsedNum) ? rawValue : parsedNum;
    } else if (field?.type === 'boolean' && typeof rawValue === 'string') {
      normalized = rawValue.toLowerCase() === 'true';
    }

    updateActiveSheet((sheet) => ({
      ...sheet,
      rows: sheet.rows.map((r) =>
        r.id === rowId ? { ...r, cells: { ...r.cells, [fieldId]: normalized } } : r
      ),
    }));
  };

  const buildDefaultRowCells = (): Record<string, CellValue> => {
    const initialCells: Record<string, CellValue> = {};
    activeSheet.fields.forEach((f) => {
      if (f.type === 'boolean') {
        initialCells[f.id] = f.defaultValue === 'true';
      } else if (f.type === 'number' && f.defaultValue !== '') {
        initialCells[f.id] = Number(f.defaultValue) || 0;
      } else {
        initialCells[f.id] = f.defaultValue || '';
      }
    });
    return initialCells;
  };

  const handleOpenAddRowModal = () => {
    setNewRowDraft(buildDefaultRowCells());
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
    insertNewRowWithCells(buildDefaultRowCells());
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
      setNewRowDraft(buildDefaultRowCells());
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
    const colLetter = getExcelColumnLetter(activeSheet.fields.length);
    const isHeaderless = activeSheet.headerRowIndex === -1;
    const label = isHeaderless ? colLetter : `Column ${colLetter}`;
    const newField: FieldConfig = {
      id: `fld_${Date.now().toString(36)}`,
      key: `col_${colLetter.toLowerCase()}`,
      label,
      sourceHeader: isHeaderless ? `(No Header — Col ${colLetter})` : label,
      columnLetter: colLetter,
      type: 'text',
      visible: true,
      width: 170,
      defaultValue: '',
      selectOptions: ['Option A', 'Option B'],
      validation: {},
    };

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
    const newId = `sheet_${Date.now().toString(36)}`;
    const colAId = `${newId}_c0_col_a`;
    const colBId = `${newId}_c1_col_b`;
    const colCId = `${newId}_c2_col_c`;
    const newSheet: SheetData = {
      id: newId,
      name: trimmed,
      description: '',
      hiddenState: 0,
      headerRowIndex: -1,
      fields: [
        {
          id: colAId,
          key: 'col_a',
          label: 'A',
          sourceHeader: '(No Header — Col A)',
          columnLetter: 'A',
          type: 'text',
          visible: true,
          width: 170,
          defaultValue: '',
          validation: {},
        },
        {
          id: colBId,
          key: 'col_b',
          label: 'B',
          sourceHeader: '(No Header — Col B)',
          columnLetter: 'B',
          type: 'text',
          visible: true,
          width: 170,
          defaultValue: '',
          validation: {},
        },
        {
          id: colCId,
          key: 'col_c',
          label: 'C',
          sourceHeader: '(No Header — Col C)',
          columnLetter: 'C',
          type: 'text',
          visible: true,
          width: 170,
          defaultValue: '',
          validation: {},
        },
      ],
      rows: [
        {
          id: `row_${Date.now().toString(36)}`,
          excelRowNumber: 1,
          cells: {
            [colAId]: '',
            [colBId]: '',
            [colCId]: '',
          },
        },
      ],
    };

    setProject((prev) => ({
      ...prev,
      sheets: [...prev.sheets, newSheet],
    }));
    setActiveSheetId(newId);
    setSelectedFieldId(colAId);
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
      <header className="top-navbar">
        <div className="brand-cluster">
          <span className="brand-title">Tabula Loom</span>
          <div className="divider-vertical" />
          <input
            type="text"
            className="workbook-title-input"
            value={project.fileName}
            onChange={(e) => setProject((p) => ({ ...p, fileName: e.target.value }))}
            aria-label="Workbook file name"
            title="Click to rename workbook"
          />
          {totalErrorCount > 0 && (
            <span
              className="status-pill invalid"
              title="Click to jump to the first cell that needs attention"
              onClick={() => {
                const firstErr = Object.values(activeSheetErrors)[0];
                if (firstErr) {
                  setSelectedCell({ rowId: firstErr.rowId, fieldId: firstErr.fieldId });
                  setSelectedFieldId(firstErr.fieldId);
                }
              }}
            >
              <IconWarning size={14} /> {totalErrorCount}{' '}
              {totalErrorCount === 1 ? 'cell needs' : 'cells need'} attention
            </span>
          )}
          {saveToastMessage ? (
            <span className="status-pill info">{saveToastMessage}</span>
          ) : draftState === 'saved' ? (
            <span className="status-pill info" title="Unsaved changes are kept in this browser">
              Draft saved locally
            </span>
          ) : null}
        </div>

        {/* Center View Switcher */}
        <div className="mode-switcher" role="tablist" aria-label="Workspace view mode">
          <button
            type="button"
            role="tab"
            aria-selected={workspaceMode === 'grid'}
            className={`mode-btn ${workspaceMode === 'grid' ? 'active' : ''}`}
            onClick={() => setWorkspaceMode('grid')}
          >
            <IconTable size={15} />
            Spreadsheet
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={workspaceMode === 'schema'}
            className={`mode-btn ${workspaceMode === 'schema' ? 'active' : ''}`}
            onClick={() => setWorkspaceMode('schema')}
          >
            <IconSliders size={15} />
            All Column Settings ({activeSheet.fields.length})
          </button>
        </div>

        {/* Right File & Zoom Actions */}
        <div className="top-actions">
          {/* Text Size / Zoom Control for Older Users */}
          <select
            value={uiScale}
            onChange={(e) => setUiScale(parseFloat(e.target.value))}
            title="Adjust spreadsheet text and row size"
            aria-label="Text size"
            style={{
              flexShrink: 0,
              whiteSpace: 'nowrap',
              padding: '5px 28px 5px 10px',
              fontSize: '13px',
              fontFamily: 'var(--font-mono)',
              color: 'var(--text-primary)',
              backgroundColor: 'var(--bg-surface)',
              border: '1px solid var(--border-hairline)',
              borderRadius: '5px',
              cursor: 'pointer',
            }}
          >
            {[0.7, 0.8, 0.9, 1.0, 1.15, 1.3, 1.5, 1.75, 2.0].map((v) => (
              <option key={v} value={v}>
                {Math.round(v * 100)}%
              </option>
            ))}
          </select>

          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            style={{ display: 'none' }}
            onChange={handleFileUpload}
          />
          <button
            type="button"
            className="btn"
            onClick={() => fileInputRef.current?.click()}
            title="Open an Excel (.xlsx, .xls) or CSV file from your computer"
          >
            <IconUpload size={15} />
            Open Excel
          </button>
          <button
            type="button"
            className="btn"
            onClick={() => {
              fetchBackendStatusAndWorkbooks();
              dbDrawerDialogRef.current?.showModal();
            }}
            title="Open or manage saved files in database"
          >
            <IconDatabase size={15} />
            Saved Files ({savedWorkbooks.length})
          </button>
          <button
            type="button"
            className="btn btn-primary"
            disabled={isSavingToDb}
            onClick={handleSaveToPostgres}
            title="Save your work to the database to continue later"
          >
            {isSavingToDb ? 'Saving...' : 'Save'}
          </button>
          <button
            type="button"
            className="btn"
            onClick={() => exportDialogRef.current?.showModal()}
          >
            <IconDownload size={15} />
            Download Excel
          </button>
        </div>
      </header>

      {/* Restored Local Draft Banner */}
      {restoredDraftAt && (
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '10px 20px',
            backgroundColor: 'var(--pastel-blue-bg)',
            borderBottom: '1px solid var(--border-hairline)',
            color: 'var(--pastel-blue-fg)',
            fontSize: '14px',
            flexShrink: 0,
            gap: '12px',
          }}
        >
          <span>
            Restored your unsaved draft from {new Date(restoredDraftAt).toLocaleString()}. Press
            Save to keep it.
          </span>
          <div style={{ display: 'flex', gap: '8px', flexShrink: 0 }}>
            <button type="button" className="btn btn-sm" onClick={handleDiscardDraft}>
              Discard draft
            </button>
            <button type="button" className="btn btn-sm" onClick={() => setRestoredDraftAt(null)}>
              Dismiss
            </button>
          </div>
        </div>
      )}

      {/* Hidden Sheets Detection Notice Banner */}
      {hiddenSheets.length > 0 && !dismissedHiddenNotice && (
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '10px 20px',
            backgroundColor: 'var(--pastel-yellow-bg)',
            borderBottom: '1px solid #E5D4A0',
            color: 'var(--pastel-yellow-fg)',
            fontSize: '14px',
            flexShrink: 0,
            gap: '12px',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
            <IconWarning size={16} />
            <span>
              <strong>
                This Excel file contains {hiddenSheets.length} hidden{' '}
                {hiddenSheets.length === 1 ? 'sheet' : 'sheets'}:
              </strong>{' '}
              {hiddenSheets.map((hs) => hs.name).join(', ')}
            </span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
            <button
              type="button"
              className="btn btn-sm"
              style={{
                backgroundColor: '#FFFFFF',
                borderColor: '#D6C080',
                color: '#6E4800',
              }}
              onClick={() => setShowHiddenSheets((prev) => !prev)}
            >
              <IconEye visible={showHiddenSheets} size={14} />
              {showHiddenSheets ? 'Hide Hidden Sheets' : 'Show Hidden Sheets'}
            </button>
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              style={{ color: '#6E4800' }}
              onClick={() => setDismissedHiddenNotice(true)}
            >
              Dismiss
            </button>
          </div>
        </div>
      )}

      {/* Main Workspace Body */}
      <div className="workspace-body">
        {workspaceMode === 'grid' ? (
          <>
            {/* Full-Width Spreadsheet Stage */}
            <main className="primary-stage">
              {/* Single Clean Excel Formula & Action Bar */}
              <div className="formula-bar">
                {activeCellInfo ? (
                  <>
                    {activeCellInfo.field.type === 'formula' ? (
                      <input
                        type="text"
                        className="formula-bar-input"
                        value={activeCellInfo.field.formulaExpression || ''}
                        placeholder="Enter formula, e.g. {col_a} * {col_b}"
                        onChange={(e) =>
                          updateFieldConfig(activeCellInfo.field.id, {
                            formulaExpression: e.target.value,
                          })
                        }
                      />
                    ) : (
                      <input
                        type="text"
                        className="formula-bar-input"
                        value={String(activeCellInfo.rawValue ?? '')}
                        placeholder={`Type value for ${activeCellInfo.field.label}...`}
                        onChange={(e) =>
                          updateCellValue(
                            activeCellInfo.row.id,
                            activeCellInfo.field.id,
                            e.target.value
                          )
                        }
                      />
                    )}
                    {activeCellInfo.err && (
                      <span className="status-pill invalid">
                        <IconWarning size={13} /> {activeCellInfo.err.message}
                      </span>
                    )}
                  </>
                ) : (
                  <span className="formula-hint">Click any cell to edit its value</span>
                )}

                {/* Right-side simple sheet controls */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
                  <div className="search-box">
                    <IconSearch size={14} />
                    <input
                      type="text"
                      placeholder="Search in sheet..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                    />
                  </div>

                  <button type="button" className="btn btn-sm" onClick={handleOpenAddRowModal}>
                    <IconPlus size={14} /> Add Row
                  </button>
                  <button type="button" className="btn btn-sm" onClick={handleAddField}>
                    <IconPlus size={14} /> Add Column
                  </button>
                  <button
                    type="button"
                    className={`btn btn-sm ${isInspectorOpen ? 'btn-primary' : ''}`}
                    onClick={() => setIsInspectorOpen((prev) => !prev)}
                    title="Open or close settings for the selected column"
                  >
                    <IconSliders size={14} />
                    {isInspectorOpen ? 'Hide Column Settings' : 'Column Settings'}
                  </button>
                </div>
              </div>

              {/* Scrollable Multi-Sheet Grid */}
              <div className="grid-scroll-container">
                <table
                  className="sheet-table"
                  style={{ width: 44 + visibleFields.reduce((s, f, idx) => s + getColWidth(f, idx), 0) }}
                >
                  <colgroup>
                    <col style={{ width: 44 }} />
                    {visibleFields.map((f, idx) => (
                      <col key={f.id} style={{ width: getColWidth(f, idx) }} />
                    ))}
                  </colgroup>
                  <thead>
                    {/* Classic Excel Row 1: Column Alphabet Letters (A, B, C, D...) */}
                    <tr className="excel-letter-row">
                      <th className="th-row-index" />
                      {visibleFields.map((field, colIdx) => {
                        const isSelectedCol = field.id === selectedFieldId;
                        const colLetter = field.columnLetter || getExcelColumnLetter(colIdx);
                        return (
                          <th
                            key={`letter_${field.id}`}
                            className={isSelectedCol ? 'selected-column-letter' : ''}
                            onClick={() => setSelectedFieldId(field.id)}
                            onDoubleClick={() => {
                              setSelectedFieldId(field.id);
                              setIsInspectorOpen(true);
                            }}
                            title="Click to select column, double-click to configure column settings"
                          >
                            {colLetter}
                            <div
                              className="col-resize-handle"
                              title="Drag to resize column"
                              onMouseDown={(e) => startColumnResize(e, field, colIdx)}
                              onClick={(e) => e.stopPropagation()}
                              onDoubleClick={(e) => e.stopPropagation()}
                            />
                          </th>
                        );
                      })}
                    </tr>

                    {/* Row 2: Clean Human-Readable Column Names (only shown when sheet has headers or custom column names) */}
                    {!isHeaderlessSheet && (
                      <tr className="column-title-row">
                        <th className="th-row-index">#</th>
                        {visibleFields.map((field, colIdx) => {
                          const isSelectedCol = field.id === selectedFieldId;
                          return (
                            <th
                              key={field.id}
                              className={isSelectedCol ? 'selected-column' : ''}
                              onClick={() => setSelectedFieldId(field.id)}
                              onDoubleClick={() => {
                                setSelectedFieldId(field.id);
                                setIsInspectorOpen(true);
                              }}
                              title="Double-click to open Column Settings"
                            >
                              <div className="th-header-single-line">
                                <div className="th-label-group">
                                  <span className="th-label">{field.label}</span>
                                  {field.validation.required && (
                                    <span className="req-dot" title="Required column">
                                      *
                                    </span>
                                  )}
                                </div>
                                <FieldTypeGlyph type={field.type} />
                              </div>
                              <div
                                className="col-resize-handle"
                                title="Drag to resize column"
                                onMouseDown={(e) => startColumnResize(e, field, colIdx)}
                                onClick={(e) => e.stopPropagation()}
                                onDoubleClick={(e) => e.stopPropagation()}
                              />
                            </th>
                          );
                        })}
                      </tr>
                    )}
                  </thead>
                  <tbody>
                    {filteredRows.map((row, rowIndex) => {
                      const displayRowNumber = row.excelRowNumber ?? rowIndex + 1;
                      return (
                        <tr key={row.id}>
                          <td className="td-row-index" title={`Excel Row ${displayRowNumber}`}>
                            <div className="row-index-inner">
                              <span>{displayRowNumber}</span>
                              <button
                                type="button"
                                className="row-delete-btn"
                                title="Delete this row"
                                onClick={() => handleDeleteRow(row.id)}
                              >
                                <IconTrash size={13} />
                              </button>
                            </div>
                          </td>
                          {visibleFields.map((field) => {
                            const cellKey = `${row.id}:${field.id}`;
                            const cellError = activeSheetErrors[cellKey];
                            const isSelected =
                              selectedCell?.rowId === row.id && selectedCell?.fieldId === field.id;
                            const isEditing =
                              editingCell?.rowId === row.id && editingCell?.fieldId === field.id;
                            const rawVal = row.cells[field.id];
                            const formatted = formatCellValue(
                              rawVal,
                              field,
                              row,
                              activeSheet.fields
                            );

                            const lookupChoices =
                              field.type === 'lookup' ? getLookupOptions(field) : [];
                            const matchedLookup =
                              field.type === 'lookup'
                                ? lookupChoices.find((c) => c.keyVal === String(rawVal ?? ''))
                                : undefined;

                            return (
                              <td
                                key={field.id}
                                className={[
                                  'grid-cell',
                                  isSelected ? 'cell-selected' : '',
                                  cellError ? 'cell-error' : '',
                                  field.type === 'formula' ? 'cell-formula' : '',
                                ]
                                  .filter(Boolean)
                                  .join(' ')}
                                onClick={() => {
                                  setSelectedCell({ rowId: row.id, fieldId: field.id });
                                  setSelectedFieldId(field.id);
                                }}
                                onDoubleClick={() => {
                                  if (field.type !== 'formula') {
                                    setEditingCell({ rowId: row.id, fieldId: field.id });
                                  }
                                }}
                              >
                                {isEditing ? (
                                  field.type === 'select' ? (
                                    <select
                                      autoFocus
                                      className="cell-inline-select"
                                      value={String(rawVal ?? '')}
                                      onChange={(e) => {
                                        updateCellValue(row.id, field.id, e.target.value);
                                        setEditingCell(null);
                                      }}
                                      onBlur={() => setEditingCell(null)}
                                    >
                                      <option value="">— Select —</option>
                                      {(field.selectOptions || []).map((opt) => (
                                        <option key={opt} value={opt}>
                                          {opt}
                                        </option>
                                      ))}
                                    </select>
                                  ) : field.type === 'lookup' ? (
                                    <select
                                      autoFocus
                                      className="cell-inline-select"
                                      value={String(rawVal ?? '')}
                                      onChange={(e) => {
                                        updateCellValue(row.id, field.id, e.target.value);
                                        setEditingCell(null);
                                      }}
                                      onBlur={() => setEditingCell(null)}
                                    >
                                      <option value="">— Select Linked Value —</option>
                                      {lookupChoices.map((item) => (
                                        <option key={item.keyVal} value={item.keyVal}>
                                          {item.keyVal}
                                          {item.displayVal ? ` — ${item.displayVal}` : ''}
                                        </option>
                                      ))}
                                    </select>
                                  ) : field.type === 'boolean' ? (
                                    <select
                                      autoFocus
                                      className="cell-inline-select"
                                      value={String(Boolean(rawVal))}
                                      onChange={(e) => {
                                        updateCellValue(
                                          row.id,
                                          field.id,
                                          e.target.value === 'true'
                                        );
                                        setEditingCell(null);
                                      }}
                                      onBlur={() => setEditingCell(null)}
                                    >
                                      <option value="true">Yes / True</option>
                                      <option value="false">No / False</option>
                                    </select>
                                  ) : (
                                    <input
                                      autoFocus
                                      type={field.type === 'date' ? 'date' : 'text'}
                                      className={`cell-inline-input ${
                                        field.type === 'number' || field.type === 'date'
                                          ? 'mono'
                                          : ''
                                      }`}
                                      defaultValue={String(rawVal ?? '')}
                                      onBlur={(e) => {
                                        updateCellValue(row.id, field.id, e.target.value);
                                        setEditingCell(null);
                                      }}
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter') {
                                          updateCellValue(
                                            row.id,
                                            field.id,
                                            (e.target as HTMLInputElement).value
                                          );
                                          setEditingCell(null);
                                        } else if (e.key === 'Escape') {
                                          setEditingCell(null);
                                        }
                                      }}
                                    />
                                  )
                                ) : (
                                  <div
                                    className={[
                                      'cell-display',
                                      field.type === 'number' || field.type === 'formula'
                                        ? 'align-right'
                                        : '',
                                      field.type === 'date' ? 'mono' : '',
                                    ]
                                      .filter(Boolean)
                                      .join(' ')}
                                    title={
                                      cellError
                                        ? cellError.message
                                        : 'Double-click to edit cell'
                                    }
                                  >
                                    {field.type === 'select' && formatted ? (
                                      <span className="cell-tag">{formatted}</span>
                                    ) : field.type === 'lookup' && formatted ? (
                                      <span
                                        style={{
                                          display: 'inline-flex',
                                          alignItems: 'center',
                                          minWidth: 0,
                                          overflow: 'hidden',
                                        }}
                                      >
                                        <span className="cell-lookup-chip">
                                          <IconLink size={12} />
                                          {formatted}
                                        </span>
                                        {matchedLookup?.displayVal && (
                                          <span className="cell-lookup-label">
                                            {matchedLookup.displayVal}
                                          </span>
                                        )}
                                      </span>
                                    ) : field.type === 'boolean' ? (
                                      <input
                                        type="checkbox"
                                        checked={
                                          rawVal === true ||
                                          String(rawVal).toLowerCase() === 'true'
                                        }
                                        onChange={(e) =>
                                          updateCellValue(row.id, field.id, e.target.checked)
                                        }
                                        style={{
                                          cursor: 'pointer',
                                          accentColor: '#111111',
                                          width: '16px',
                                          height: '16px',
                                        }}
                                      />
                                    ) : (
                                      <span
                                        style={{
                                          overflow: 'hidden',
                                          textOverflow: 'ellipsis',
                                          width: '100%',
                                        }}
                                      >
                                        {formatted}
                                      </span>
                                    )}

                                    {cellError && (
                                      <span
                                        className="cell-error-indicator"
                                        title={cellError.message}
                                      >
                                        <IconWarning size={15} />
                                      </span>
                                    )}
                                  </div>
                                )}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </main>

            {/* Collapsible Right Column Settings Drawer (Only shown when user clicks "Column Settings") */}
            {isInspectorOpen && selectedField && (
              <aside className="field-inspector" aria-label="Column Settings">
                <div className="inspector-header">
                  <div>
                    <div className="section-heading">
                      Column {selectedField.columnLetter || '—'} Settings
                    </div>
                    <h2 className="inspector-title">
                      {selectedField.label === selectedField.columnLetter
                        ? `Column ${selectedField.columnLetter}`
                        : selectedField.label}
                    </h2>
                  </div>
                  <button
                    type="button"
                    className="btn btn-sm"
                    onClick={() => setIsInspectorOpen(false)}
                  >
                    Close
                  </button>
                </div>

                {/* Essential Settings in Plain Language */}
                <div className="inspector-section">
                  <div className="form-field">
                    <label className="form-label">Column Name</label>
                    <input
                      type="text"
                      className="input-control"
                      value={selectedField.label}
                      onChange={(e) =>
                        updateFieldConfig(selectedField.id, {
                          label: e.target.value,
                        })
                      }
                    />
                  </div>

                  <div className="form-field">
                    <label className="form-label">Column Data Type</label>
                    <select
                      className="select-control"
                      value={selectedField.type}
                      onChange={(e) => {
                        const nextType = e.target.value as FieldDataType;
                        const patch: Partial<FieldConfig> = { type: nextType };
                        if (nextType === 'lookup' && !selectedField.lookup) {
                          const otherSheet =
                            project.sheets.find((s) => s.id !== activeSheet.id) ||
                            project.sheets[0];
                          patch.lookup = {
                            targetSheetId: otherSheet.id,
                            targetFieldId: otherSheet.fields[0]?.id || '',
                            displayFieldId: otherSheet.fields[1]?.id || otherSheet.fields[0]?.id,
                          };
                        }
                        updateFieldConfig(selectedField.id, patch);
                      }}
                    >
                      <option value="text">Text</option>
                      <option value="number">Number / Currency</option>
                      <option value="date">Date</option>
                      <option value="select">Dropdown List (Options)</option>
                      <option value="boolean">Checkbox (Yes / No)</option>
                      <option value="formula">Formula Calculation</option>
                      <option value="lookup">Link to Another Sheet</option>
                    </select>
                  </div>

                  {(selectedField.type === 'number' || selectedField.type === 'formula') && (
                    <div className="form-row-2">
                      <div className="form-field">
                        <label className="form-label">Number Style</label>
                        <select
                          className="select-control"
                          value={selectedField.numberFormat || 'raw'}
                          onChange={(e) =>
                            updateFieldConfig(selectedField.id, {
                              numberFormat: e.target.value as NumberFormatType,
                            })
                          }
                        >
                          <option value="raw">Plain Number</option>
                          <option value="integer">Whole Number (1,000)</option>
                          <option value="decimal">Two Decimals (1,000.00)</option>
                          <option value="currency">Currency ($1,000.00)</option>
                          <option value="percent">Percent (10.0%)</option>
                        </select>
                      </div>
                      <div className="form-field">
                        <label className="form-label">Symbol</label>
                        <input
                          type="text"
                          className="input-control mono"
                          value={selectedField.currencySymbol || '$'}
                          onChange={(e) =>
                            updateFieldConfig(selectedField.id, {
                              currencySymbol: e.target.value,
                            })
                          }
                        />
                      </div>
                    </div>
                  )}

                  {selectedField.type === 'select' && (
                    <div className="form-field">
                      <label className="form-label">Dropdown Choices (comma-separated)</label>
                      <textarea
                        className="textarea-control"
                        value={(selectedField.selectOptions || []).join(', ')}
                        onChange={(e) =>
                          updateFieldConfig(selectedField.id, {
                            selectOptions: e.target.value
                              .split(',')
                              .map((s) => s.trim())
                              .filter(Boolean),
                          })
                        }
                      />
                    </div>
                  )}

                  {selectedField.type === 'formula' && (
                    <div className="form-field">
                      <label className="form-label">Formula</label>
                      <input
                        type="text"
                        className="input-control mono"
                        placeholder="e.g. {col_a} * {col_b}"
                        value={selectedField.formulaExpression || ''}
                        onChange={(e) =>
                          updateFieldConfig(selectedField.id, {
                            formulaExpression: e.target.value,
                          })
                        }
                      />
                    </div>
                  )}

                  {selectedField.type === 'lookup' && (
                    <>
                      <div className="form-field">
                        <label className="form-label">Link to Sheet</label>
                        <select
                          className="select-control"
                          value={selectedField.lookup?.targetSheetId || ''}
                          onChange={(e) => {
                            const targetSheet = project.sheets.find(
                              (s) => s.id === e.target.value
                            );
                            if (!targetSheet) return;
                            updateFieldConfig(selectedField.id, {
                              lookup: {
                                targetSheetId: targetSheet.id,
                                targetFieldId: targetSheet.fields[0]?.id || '',
                                displayFieldId:
                                  targetSheet.fields[1]?.id || targetSheet.fields[0]?.id,
                              },
                            });
                          }}
                        >
                          {project.sheets.map((s) => (
                            <option key={s.id} value={s.id}>
                              {s.name}
                            </option>
                          ))}
                        </select>
                      </div>

                      {(() => {
                        const targetSheet = project.sheets.find(
                          (s) => s.id === selectedField.lookup?.targetSheetId
                        );
                        if (!targetSheet) return null;
                        return (
                          <div className="form-row-2">
                            <div className="form-field">
                              <label className="form-label">Match Column</label>
                              <select
                                className="select-control"
                                value={selectedField.lookup?.targetFieldId || ''}
                                onChange={(e) =>
                                  updateFieldConfig(selectedField.id, {
                                    lookup: {
                                      ...selectedField.lookup!,
                                      targetFieldId: e.target.value,
                                    },
                                  })
                                }
                              >
                                {targetSheet.fields.map((tf) => (
                                  <option key={tf.id} value={tf.id}>
                                    {tf.label}
                                  </option>
                                ))}
                              </select>
                            </div>
                            <div className="form-field">
                              <label className="form-label">Display Column</label>
                              <select
                                className="select-control"
                                value={selectedField.lookup?.displayFieldId || ''}
                                onChange={(e) =>
                                  updateFieldConfig(selectedField.id, {
                                    lookup: {
                                      ...selectedField.lookup!,
                                      displayFieldId: e.target.value,
                                    },
                                  })
                                }
                              >
                                {targetSheet.fields.map((tf) => (
                                  <option key={tf.id} value={tf.id}>
                                    {tf.label}
                                  </option>
                                ))}
                              </select>
                            </div>
                          </div>
                        );
                      })()}
                    </>
                  )}
                </div>

                {/* Simple Validation Checkboxes */}
                <div className="inspector-section">
                  <div className="section-heading">Data Rules</div>

                  <div className="form-row-2">
                    <label className="checkbox-row">
                      <span>Required</span>
                      <input
                        type="checkbox"
                        checked={Boolean(selectedField.validation.required)}
                        onChange={(e) =>
                          updateFieldConfig(selectedField.id, {
                            validation: {
                              ...selectedField.validation,
                              required: e.target.checked,
                            },
                          })
                        }
                      />
                    </label>
                    <label className="checkbox-row">
                      <span>No Duplicates</span>
                      <input
                        type="checkbox"
                        checked={Boolean(selectedField.validation.unique)}
                        onChange={(e) =>
                          updateFieldConfig(selectedField.id, {
                            validation: {
                              ...selectedField.validation,
                              unique: e.target.checked,
                            },
                          })
                        }
                      />
                    </label>
                  </div>

                  {selectedField.type === 'number' && (
                    <div className="form-row-2">
                      <div className="form-field">
                        <label className="form-label">Minimum</label>
                        <input
                          type="number"
                          className="input-control mono"
                          placeholder="No min"
                          value={selectedField.validation.min ?? ''}
                          onChange={(e) =>
                            updateFieldConfig(selectedField.id, {
                              validation: {
                                ...selectedField.validation,
                                min:
                                  e.target.value === '' ? undefined : Number(e.target.value),
                              },
                            })
                          }
                        />
                      </div>
                      <div className="form-field">
                        <label className="form-label">Maximum</label>
                        <input
                          type="number"
                          className="input-control mono"
                          placeholder="No max"
                          value={selectedField.validation.max ?? ''}
                          onChange={(e) =>
                            updateFieldConfig(selectedField.id, {
                              validation: {
                                ...selectedField.validation,
                                max:
                                  e.target.value === '' ? undefined : Number(e.target.value),
                              },
                            })
                          }
                        />
                      </div>
                    </div>
                  )}

                  <div className="form-field">
                    <label className="form-label">Default Value for New Rows</label>
                    <input
                      type="text"
                      className="input-control"
                      placeholder="Blank"
                      value={selectedField.defaultValue}
                      onChange={(e) =>
                        updateFieldConfig(selectedField.id, { defaultValue: e.target.value })
                      }
                    />
                  </div>
                </div>

                {/* Column Position & Visibility */}
                <div className="inspector-section" style={{ borderBottom: 'none' }}>
                  <div className="section-heading">Column Layout</div>
                  <div className="form-row-2">
                    <button
                      type="button"
                      className="btn"
                      onClick={() => handleMoveField(selectedField.id, -1)}
                    >
                      ← Move Left
                    </button>
                    <button
                      type="button"
                      className="btn"
                      onClick={() => handleMoveField(selectedField.id, 1)}
                    >
                      Move Right →
                    </button>
                  </div>

                  <div className="form-row-2">
                    <label className="checkbox-row">
                      <span>Show Column</span>
                      <input
                        type="checkbox"
                        checked={selectedField.visible}
                        onChange={(e) =>
                          updateFieldConfig(selectedField.id, { visible: e.target.checked })
                        }
                      />
                    </label>
                    <div className="form-field">
                      <input
                        type="number"
                        className="input-control mono"
                        min={100}
                        max={450}
                        value={selectedField.width}
                        onChange={(e) =>
                          updateFieldConfig(selectedField.id, {
                            width: Number(e.target.value) || 160,
                          })
                        }
                        title="Column width in pixels"
                      />
                    </div>
                  </div>

                  <button
                    type="button"
                    className="btn btn-danger"
                    style={{ marginTop: '8px' }}
                    disabled={activeSheet.fields.length <= 1}
                    onClick={() => handleDeleteField(selectedField.id)}
                  >
                    <IconTrash size={14} /> Delete Column
                  </button>
                </div>
              </aside>
            )}
          </>
        ) : (
          /* All Column Settings View */
          <main className="schema-view-container">
            <div className="schema-view-inner">
              <div className="schema-hero-card">
                <div>
                  <div className="section-heading">Sheet Column Settings</div>
                  <h1 className="schema-hero-title">{activeSheet.name}</h1>
                  <p style={{ color: 'var(--text-muted)', fontSize: '15px' }}>
                    Configure column names, data types, dropdown lists, and rules for this sheet.
                  </p>
                </div>
                <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
                  <div
                    style={{
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: '6px',
                      height: '38px',
                      padding: '0 12px',
                      backgroundColor: 'var(--bg-subtle)',
                      border: '1px solid var(--border-hairline)',
                      borderRadius: '6px',
                      fontSize: '13.5px',
                      color: 'var(--text-body)',
                    }}
                    title="Choose which Excel row contains the column headers for this sheet"
                  >
                    <span>Header Row:</span>
                    <select
                      value={activeSheet.headerRowIndex}
                      onChange={(e) => handleHeaderRowChange(Number(e.target.value))}
                      style={{
                        border: 'none',
                        background: 'transparent',
                        fontSize: '13.5px',
                        fontWeight: 600,
                        color: 'var(--text-primary)',
                        cursor: 'pointer',
                        paddingRight: '22px',
                      }}
                    >
                      <option value={-1}>No Header (Use A, B, C)</option>
                      {(activeSheet.rawMatrix || []).slice(0, 10).map((r, idx) => {
                        const preview = r
                          .filter((c) => String(c ?? '').trim() !== '')
                          .slice(0, 2)
                          .join(', ');
                        return (
                          <option key={idx} value={idx}>
                            Row {idx + 1} {preview ? `(${preview.slice(0, 18)})` : ''}
                          </option>
                        );
                      })}
                      {!activeSheet.rawMatrix && <option value={0}>Row 1</option>}
                    </select>
                  </div>
                  <button
                    type="button"
                    className="btn"
                    onClick={() => setShowHiddenColumns((prev) => !prev)}
                  >
                    <IconEye visible={showHiddenColumns} size={15} />
                    {showHiddenColumns ? 'Showing Hidden Columns' : 'Show Hidden Columns'}
                  </button>
                  <button type="button" className="btn" onClick={handleAddField}>
                    <IconPlus size={15} /> Add Column
                  </button>
                  <button
                    type="button"
                    className="btn btn-primary"
                    onClick={() => setWorkspaceMode('grid')}
                  >
                    <IconTable size={15} /> Back to Spreadsheet
                  </button>
                </div>
              </div>

              <div className="schema-table-card">
                <table className="schema-config-table">
                  <thead>
                    <tr>
                      <th style={{ width: '64px' }}>Show</th>
                      <th style={{ width: '64px' }}>Col</th>
                      <th style={{ width: '240px' }}>Column Name</th>
                      <th style={{ width: '200px' }}>Data Type</th>
                      <th>Options / Formula / Linked Sheet</th>
                      <th style={{ width: '150px' }}>Default Value</th>
                      <th style={{ width: '160px' }}>Rules</th>
                      <th style={{ width: '64px' }} />
                    </tr>
                  </thead>
                  <tbody>
                    {activeSheet.fields.map((field, idx) => (
                      <tr
                        key={field.id}
                        className={field.id === selectedFieldId ? 'active-schema-row' : ''}
                        onClick={() => setSelectedFieldId(field.id)}
                      >
                        <td>
                          <button
                            type="button"
                            className="btn btn-ghost btn-sm"
                            onClick={() =>
                              updateFieldConfig(field.id, { visible: !field.visible })
                            }
                            title={field.visible ? 'Visible in sheet' : 'Hidden in sheet'}
                          >
                            <IconEye visible={field.visible} size={15} />
                          </button>
                        </td>
                        <td
                          style={{
                            fontFamily: 'var(--font-mono)',
                            fontSize: '14px',
                            fontWeight: 600,
                            color: 'var(--text-muted)',
                          }}
                        >
                          {field.columnLetter || getExcelColumnLetter(idx)}
                        </td>
                        <td>
                          <input
                            type="text"
                            className="input-control"
                            value={field.label}
                            onChange={(e) =>
                              updateFieldConfig(field.id, {
                                label: e.target.value,
                                key: toMachineKey(e.target.value),
                              })
                            }
                          />
                        </td>
                        <td>
                          <select
                            className="select-control"
                            value={field.type}
                            onChange={(e) =>
                              updateFieldConfig(field.id, {
                                type: e.target.value as FieldDataType,
                              })
                            }
                          >
                            <option value="text">Text</option>
                            <option value="number">Number / Currency</option>
                            <option value="date">Date</option>
                            <option value="select">Dropdown List</option>
                            <option value="boolean">Checkbox (Yes/No)</option>
                            <option value="formula">Formula</option>
                            <option value="lookup">Link to Sheet</option>
                          </select>
                        </td>
                        <td>
                          {field.type === 'select' && (
                            <input
                              type="text"
                              className="input-control"
                              placeholder="Option 1, Option 2, Option 3"
                              value={(field.selectOptions || []).join(', ')}
                              onChange={(e) =>
                                updateFieldConfig(field.id, {
                                  selectOptions: e.target.value
                                    .split(',')
                                    .map((s) => s.trim())
                                    .filter(Boolean),
                                })
                              }
                            />
                          )}
                          {field.type === 'formula' && (
                            <input
                              type="text"
                              className="input-control mono"
                              placeholder="{col_a} * {col_b}"
                              value={field.formulaExpression || ''}
                              onChange={(e) =>
                                updateFieldConfig(field.id, {
                                  formulaExpression: e.target.value,
                                })
                              }
                            />
                          )}
                          {field.type === 'number' && (
                            <select
                              className="select-control"
                              value={field.numberFormat || 'raw'}
                              onChange={(e) =>
                                updateFieldConfig(field.id, {
                                  numberFormat: e.target.value as NumberFormatType,
                                })
                              }
                            >
                              <option value="raw">Plain Number</option>
                              <option value="integer">Whole Number (1,000)</option>
                              <option value="decimal">Two Decimals (1,000.00)</option>
                              <option value="currency">Currency ($1,000.00)</option>
                              <option value="percent">Percent (10.0%)</option>
                            </select>
                          )}
                          {field.type === 'lookup' && (
                            <div style={{ display: 'flex', gap: '8px' }}>
                              <select
                                className="select-control"
                                value={field.lookup?.targetSheetId || ''}
                                onChange={(e) => {
                                  const ts = project.sheets.find(
                                    (s) => s.id === e.target.value
                                  );
                                  if (!ts) return;
                                  updateFieldConfig(field.id, {
                                    lookup: {
                                      targetSheetId: ts.id,
                                      targetFieldId: ts.fields[0]?.id || '',
                                      displayFieldId:
                                        ts.fields[1]?.id || ts.fields[0]?.id,
                                    },
                                  });
                                }}
                              >
                                {project.sheets.map((s) => (
                                  <option key={s.id} value={s.id}>
                                    Sheet: {s.name}
                                  </option>
                                ))}
                              </select>
                              <select
                                className="select-control"
                                value={field.lookup?.targetFieldId || ''}
                                onChange={(e) =>
                                  updateFieldConfig(field.id, {
                                    lookup: {
                                      ...(field.lookup || {
                                        targetSheetId: project.sheets[0].id,
                                      }),
                                      targetFieldId: e.target.value,
                                    },
                                  })
                                }
                              >
                                {(
                                  project.sheets.find(
                                    (s) => s.id === field.lookup?.targetSheetId
                                  )?.fields || []
                                ).map((tf) => (
                                  <option key={tf.id} value={tf.id}>
                                    Column: {tf.label}
                                  </option>
                                ))}
                              </select>
                            </div>
                          )}
                          {(field.type === 'text' ||
                            field.type === 'date' ||
                            field.type === 'boolean') && (
                            <span className="form-hint">Standard {field.type} column</span>
                          )}
                        </td>
                        <td>
                          <input
                            type="text"
                            className="input-control"
                            placeholder="Blank"
                            value={field.defaultValue}
                            onChange={(e) =>
                              updateFieldConfig(field.id, {
                                defaultValue: e.target.value,
                              })
                            }
                          />
                        </td>
                        <td>
                          <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                            <label
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '5px',
                                fontSize: '13.5px',
                                cursor: 'pointer',
                              }}
                            >
                              <input
                                type="checkbox"
                                checked={Boolean(field.validation.required)}
                                onChange={(e) =>
                                  updateFieldConfig(field.id, {
                                    validation: {
                                      ...field.validation,
                                      required: e.target.checked,
                                    },
                                  })
                                }
                              />
                              Required
                            </label>
                            <label
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '5px',
                                fontSize: '13.5px',
                                cursor: 'pointer',
                              }}
                            >
                              <input
                                type="checkbox"
                                checked={Boolean(field.validation.unique)}
                                onChange={(e) =>
                                  updateFieldConfig(field.id, {
                                    validation: {
                                      ...field.validation,
                                      unique: e.target.checked,
                                    },
                                  })
                                }
                              />
                              Unique
                            </label>
                          </div>
                        </td>
                        <td>
                          <button
                            type="button"
                            className="btn btn-ghost btn-sm"
                            disabled={activeSheet.fields.length <= 1}
                            onClick={() => handleDeleteField(field.id)}
                            title="Delete column"
                          >
                            <IconTrash size={15} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </main>
        )}
      </div>

      {/* Classic Excel Bottom Bar: Sheet Tabs on the Left, Simple Summary on the Right */}
      <footer className="excel-bottom-bar">
        <div className="sheet-tabs-bar" role="tablist" aria-label="Worksheet tabs">
          {displayedSheets.map((sheet) => {
            const sheetErrCount = Object.keys(sheetErrorsMap[sheet.id] || {}).length;
            const isActive = sheet.id === activeSheet.id;
            const isHiddenSheet = Boolean(sheet.hiddenState);
            return (
              <button
                key={sheet.id}
                type="button"
                role="tab"
                aria-selected={isActive}
                className={`sheet-tab ${isActive ? 'active' : ''}`}
                onClick={() => handleSelectSheet(sheet.id)}
                style={
                  isHiddenSheet
                    ? {
                        borderStyle: 'dashed',
                        borderColor: '#956400',
                      }
                    : undefined
                }
              >
                {isHiddenSheet && (
                  <span
                    className="status-pill warn"
                    style={{ padding: '1px 6px', fontSize: '11px' }}
                  >
                    Hidden
                  </span>
                )}
                <span>{sheet.name}</span>
                {sheetErrCount > 0 && (
                  <span
                    style={{
                      color: 'var(--pastel-red-fg)',
                      fontSize: '12.5px',
                      fontWeight: 700,
                    }}
                    title={`${sheetErrCount} cells need attention in this sheet`}
                  >
                    ({sheetErrCount})
                  </span>
                )}
              </button>
            );
          })}
          {hiddenSheets.length > 0 && dismissedHiddenNotice && (
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              onClick={() => setShowHiddenSheets((prev) => !prev)}
            >
              <IconEye visible={showHiddenSheets} size={14} />
              {showHiddenSheets
                ? `Hide Hidden (${hiddenSheets.length})`
                : `Show Hidden (${hiddenSheets.length})`}
            </button>
          )}
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => newSheetDialogRef.current?.showModal()}
            title="Add a new worksheet tab"
          >
            <IconPlus size={14} /> New Sheet
          </button>
        </div>

        <div className="bottom-right-meta">
          <span>
            {filteredRows.length} {filteredRows.length === 1 ? 'Row' : 'Rows'} •{' '}
            {visibleFields.length} {visibleFields.length === 1 ? 'Column' : 'Columns'}
          </span>
          <span>•</span>
          <span>Double-click any cell to edit</span>
        </div>
      </footer>

      {/* Modal Dialog: Saved Files in Database */}
      <dialog
        ref={dbDrawerDialogRef}
        closedby="any"
        className="editorial-dialog"
        style={{ maxWidth: '640px' }}
      >
        <div className="dialog-header">
          <h3 className="dialog-title">Saved Excel Workbooks</h3>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => dbDrawerDialogRef.current?.close()}
          >
            Close
          </button>
        </div>
        <div className="dialog-body">
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              padding: '12px 16px',
              borderRadius: '6px',
              border: '1px solid var(--border-hairline)',
              backgroundColor: 'var(--bg-subtle)',
            }}
          >
            <div>
              <span className={`status-pill ${dbStatus.connected ? 'valid' : 'warn'}`}>
                {dbStatus.connected ? 'Database Connected' : 'Database Offline'}
              </span>
              {dbStatus.error && (
                <div style={{ fontSize: '13px', color: 'var(--text-muted)', marginTop: '4px' }}>
                  {dbStatus.error}
                </div>
              )}
            </div>
          </div>

          {savedWorkbooks.length === 0 ? (
            <div
              style={{
                padding: '28px',
                textAlign: 'center',
                border: '1px dashed var(--border-strong)',
                borderRadius: '6px',
                color: 'var(--text-muted)',
                fontSize: '14.5px',
              }}
            >
              No saved files yet. Use the Save button in the top bar to store your sheets.
            </div>
          ) : (
            <div
              style={{
                display: 'flex',
                flexDirection: 'column',
                gap: '8px',
                maxHeight: '300px',
                overflowY: 'auto',
              }}
            >
              {savedWorkbooks.map((wb) => {
                const isCurrent = project.id === wb.id;
                return (
                  <div
                    key={wb.id}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      padding: '12px 16px',
                      border: '1px solid var(--border-hairline)',
                      borderRadius: '6px',
                      backgroundColor: isCurrent ? 'var(--bg-selected)' : 'var(--bg-surface)',
                    }}
                  >
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <strong style={{ fontSize: '15px', color: 'var(--text-primary)' }}>
                          {wb.fileName}
                        </strong>
                        <span className="status-pill info">
                          {wb.sheetCount} {wb.sheetCount === 1 ? 'sheet' : 'sheets'}
                        </span>
                      </div>
                      <div
                        style={{
                          fontSize: '13px',
                          color: 'var(--text-muted)',
                          marginTop: '3px',
                        }}
                      >
                        Last saved: {new Date(wb.lastModified).toLocaleString()}
                      </div>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      {wb.storageKey && (
                        <a
                          href={`/api/workbooks/${wb.id}/original-file`}
                          className="btn btn-sm"
                          style={{ textDecoration: 'none' }}
                        >
                          Original File
                        </a>
                      )}
                      <button
                        type="button"
                        className="btn btn-primary btn-sm"
                        disabled={isLoadingFromDb}
                        onClick={() => handleLoadWorkbookFromDb(wb.id)}
                      >
                        Open
                      </button>
                      <button
                        type="button"
                        className="btn btn-ghost btn-sm"
                        onClick={() => handleDeleteWorkbookFromDb(wb.id, wb.fileName)}
                        title="Delete saved workbook"
                      >
                        <IconTrash size={14} />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
        <div className="dialog-footer">
          <button
            type="button"
            className="btn btn-primary"
            onClick={() => dbDrawerDialogRef.current?.close()}
          >
            Done
          </button>
        </div>
      </dialog>

      {/* Modal Dialog: Export XLSX Options */}
      <dialog ref={exportDialogRef} closedby="any" className="editorial-dialog">
        <div className="dialog-header">
          <h3 className="dialog-title">Download Excel File (.xlsx)</h3>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => exportDialogRef.current?.close()}
          >
            Close
          </button>
        </div>
        <div className="dialog-body">
          <label className="checkbox-row">
            <div>
              <span>Include Hidden Columns</span>
              <div className="form-hint">
                Export all {activeSheet.fields.length} columns even if hidden on screen
              </div>
            </div>
            <input
              type="checkbox"
              checked={exportIncludeHidden}
              onChange={(e) => setExportIncludeHidden(e.target.checked)}
            />
          </label>

          {hiddenSheets.length > 0 && (
            <label className="checkbox-row">
              <div>
                <span>Include Hidden Sheets ({hiddenSheets.length})</span>
                <div className="form-hint">
                  Keep hidden worksheets ({hiddenSheets.map((s) => s.name).join(', ')}) in the file
                </div>
              </div>
              <input
                type="checkbox"
                checked={exportIncludeHiddenSheets}
                onChange={(e) => setExportIncludeHiddenSheets(e.target.checked)}
              />
            </label>
          )}

          <label className="checkbox-row">
            <div>
              <span>Include Column Settings Summary Sheet</span>
              <div className="form-hint">
                Adds a sheet listing your column types and rules
              </div>
            </div>
            <input
              type="checkbox"
              checked={exportIncludeSchemaTab}
              onChange={(e) => setExportIncludeSchemaTab(e.target.checked)}
            />
          </label>
        </div>
        <div className="dialog-footer">
          <button
            type="button"
            className="btn"
            onClick={() => exportDialogRef.current?.close()}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={() => {
              exportWorkbookToXlsx(project, {
                useMachineKeys: exportUseMachineKeys,
                includeHiddenFields: exportIncludeHidden,
                includeHiddenSheets: exportIncludeHiddenSheets,
                includeSchemaSheet: exportIncludeSchemaTab,
              });
              exportDialogRef.current?.close();
            }}
          >
            <IconDownload size={15} /> Download .xlsx
          </button>
        </div>
      </dialog>

      {/* Modal Dialog: Create New Sheet */}
      <dialog ref={newSheetDialogRef} closedby="any" className="editorial-dialog">
        <form onSubmit={handleCreateSheet}>
          <div className="dialog-header">
            <h3 className="dialog-title">Add New Sheet</h3>
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              onClick={() => newSheetDialogRef.current?.close()}
            >
              Close
            </button>
          </div>
          <div className="dialog-body">
            <div className="form-field">
              <label className="form-label">Sheet Name</label>
              <input
                type="text"
                className="input-control"
                placeholder="e.g. Sheet 2"
                value={newSheetName}
                onChange={(e) => setNewSheetName(e.target.value)}
                autoFocus
              />
            </div>
          </div>
          <div className="dialog-footer">
            <button
              type="button"
              className="btn"
              onClick={() => newSheetDialogRef.current?.close()}
            >
              Cancel
            </button>
            <button type="submit" className="btn btn-primary">
              Create Sheet
            </button>
          </div>
        </form>
      </dialog>

      {/* Modal Dialog: Dynamic Field-Based Add Row Form */}
      <dialog
        ref={addRowDialogRef}
        closedby="any"
        className="editorial-dialog"
        style={{ maxWidth: '640px' }}
      >
        <form onSubmit={(e) => handleSubmitAddRowForm(e, false)}>
          <div className="dialog-header">
            <div>
              <h3 className="dialog-title">Add Row to "{activeSheet.name}"</h3>
              <div style={{ fontSize: '13px', color: 'var(--text-muted)', marginTop: '2px' }}>
                Row #{(activeSheet.rows[activeSheet.rows.length - 1]?.excelRowNumber || activeSheet.rows.length) + 1} •{' '}
                {activeSheet.fields.length} configured {activeSheet.fields.length === 1 ? 'field' : 'fields'}
              </div>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <button
                type="button"
                className="btn btn-sm"
                onClick={handleAddBlankRow}
                title="Insert a blank row immediately without filling out the form"
              >
                Insert Blank Row
              </button>
              <button
                type="button"
                className="btn btn-ghost btn-sm"
                onClick={() => addRowDialogRef.current?.close()}
              >
                Close
              </button>
            </div>
          </div>

          <div
            className="dialog-body"
            style={{
              maxHeight: '62vh',
              overflowY: 'auto',
              display: 'grid',
              gridTemplateColumns: activeSheet.fields.length > 4 ? '1fr 1fr' : '1fr',
              gap: '14px 16px',
            }}
          >
            {activeSheet.fields.map((field, idx) => {
              const colLetter = field.columnLetter || getExcelColumnLetter(idx);
              const rawVal = newRowDraft[field.id] ?? '';
              const fieldErr = newRowErrors[field.id];
              const lookupChoices = field.type === 'lookup' ? getLookupOptions(field) : [];
              const formulaPreview =
                field.type === 'formula'
                  ? formatCellValue(
                      '',
                      field,
                      { id: '__preview__', cells: newRowDraft },
                      activeSheet.fields
                    )
                  : '';

              const setDraftField = (val: CellValue) => {
                setNewRowDraft((prev) => ({ ...prev, [field.id]: val }));
                if (newRowErrors[field.id]) {
                  setNewRowErrors((prev) => {
                    const next = { ...prev };
                    delete next[field.id];
                    return next;
                  });
                }
              };

              return (
                <div key={field.id} className="form-field">
                  <label className="form-label">
                    <span
                      style={{
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '6px',
                        minWidth: 0,
                      }}
                    >
                      <span
                        style={{
                          fontFamily: 'var(--font-mono)',
                          fontSize: '11.5px',
                          color: 'var(--text-muted)',
                          backgroundColor: 'var(--bg-canvas)',
                          border: '1px solid var(--border-hairline)',
                          borderRadius: '4px',
                          padding: '1px 5px',
                        }}
                      >
                        {colLetter}
                      </span>
                      <span
                        style={{
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {field.label === colLetter ? `Column ${colLetter}` : field.label}
                      </span>
                      {field.validation.required && (
                        <span className="req-dot" title="Required field">
                          *
                        </span>
                      )}
                    </span>
                    <FieldTypeGlyph type={field.type} />
                  </label>

                  {field.type === 'select' ? (
                    <select
                      className="select-control"
                      value={String(rawVal ?? '')}
                      onChange={(e) => setDraftField(e.target.value)}
                    >
                      <option value="">— Select option —</option>
                      {(field.selectOptions || []).map((opt) => (
                        <option key={opt} value={opt}>
                          {opt}
                        </option>
                      ))}
                    </select>
                  ) : field.type === 'lookup' ? (
                    <select
                      className="select-control"
                      value={String(rawVal ?? '')}
                      onChange={(e) => setDraftField(e.target.value)}
                    >
                      <option value="">— Select linked record —</option>
                      {lookupChoices.map((item) => (
                        <option key={item.keyVal} value={item.keyVal}>
                          {item.keyVal}
                          {item.displayVal ? ` — ${item.displayVal}` : ''}
                        </option>
                      ))}
                    </select>
                  ) : field.type === 'boolean' ? (
                    <label className="checkbox-row" style={{ height: '38px' }}>
                      <span>
                        {rawVal === true || String(rawVal).toLowerCase() === 'true'
                          ? 'Yes / True'
                          : 'No / False'}
                      </span>
                      <input
                        type="checkbox"
                        checked={rawVal === true || String(rawVal).toLowerCase() === 'true'}
                        onChange={(e) => setDraftField(e.target.checked)}
                        style={{ accentColor: '#111111', width: '16px', height: '16px' }}
                      />
                    </label>
                  ) : field.type === 'formula' ? (
                    <input
                      type="text"
                      className="input-control mono"
                      value={formulaPreview || '0'}
                      readOnly
                      disabled
                      title={`Calculated automatically: ${field.formulaExpression || 'Formula'}`}
                      style={{
                        backgroundColor: '#FCFBF7',
                        color: 'var(--text-muted)',
                        cursor: 'not-allowed',
                      }}
                    />
                  ) : (
                    <input
                      type={
                        field.type === 'date'
                          ? 'date'
                          : field.type === 'number'
                          ? 'number'
                          : 'text'
                      }
                      step={field.type === 'number' ? 'any' : undefined}
                      className={`input-control ${
                        field.type === 'number' || field.type === 'date' ? 'mono' : ''
                      }`}
                      placeholder={
                        field.type === 'number'
                          ? field.numberFormat === 'currency'
                            ? `${field.currencySymbol || '$'}0.00`
                            : 'Enter number...'
                          : `Enter ${field.label}...`
                      }
                      value={String(rawVal ?? '')}
                      onChange={(e) => setDraftField(e.target.value)}
                      autoFocus={idx === 0}
                    />
                  )}

                  {fieldErr ? (
                    <div
                      style={{
                        fontSize: '12.5px',
                        color: 'var(--pastel-red-fg)',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '4px',
                      }}
                    >
                      <IconWarning size={13} /> {fieldErr}
                    </div>
                  ) : field.description ? (
                    <div className="form-hint">{field.description}</div>
                  ) : null}
                </div>
              );
            })}
          </div>

          <div className="dialog-footer">
            <button
              type="button"
              className="btn"
              onClick={() => addRowDialogRef.current?.close()}
            >
              Cancel
            </button>
            <button
              type="button"
              className="btn"
              onClick={(e) => handleSubmitAddRowForm(e, true)}
              title="Insert this row and reset the form to add another row"
            >
              <IconPlus size={14} /> Add &amp; Add Another
            </button>
            <button type="submit" className="btn btn-primary">
              <IconCheck size={14} /> Add Row
            </button>
          </div>
        </form>
      </dialog>
    </div>
  );
}
