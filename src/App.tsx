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
  const [newSheetName, setNewSheetName] = useState<string>('');

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
        error: 'Backend API server is not running on port 3001.',
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
        showToast(data.error || 'Could not save to PostgreSQL');
        return;
      }
      setProject((prev) => ({
        ...prev,
        id: data.id,
        lastModified: data.lastModified,
        hasOriginalBinary: prev.hasOriginalBinary || Boolean(rawUploadedBase64),
      }));
      setRawUploadedBase64(null);
      await fetchBackendStatusAndWorkbooks();
      showToast(`Saved "${project.fileName}" (${project.sheets.length} sheets) to PostgreSQL`);
    } catch {
      dbDrawerDialogRef.current?.showModal();
      showToast('Failed to reach PostgreSQL API');
    } finally {
      setIsSavingToDb(false);
    }
  };

  const handleLoadWorkbookFromDb = async (workbookId: string) => {
    setIsLoadingFromDb(true);
    try {
      const res = await fetch(`/api/workbooks/${workbookId}`);
      if (!res.ok) {
        showToast('Failed to load workbook from PostgreSQL');
        return;
      }
      const loaded: WorkbookProject = await res.json();
      if (loaded.sheets.length > 0) {
        setProject(loaded);
        setRawUploadedBase64(null);
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
        showToast(`Loaded "${loaded.fileName}" from PostgreSQL`);
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
      showToast(`Deleted "${fileName}" from PostgreSQL`);
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
      // Prefer selecting the first visible sheet by default
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

  const handleAddRow = () => {
    const newRowId = `row_${Date.now().toString(36)}`;
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

    updateActiveSheet((sheet) => ({
      ...sheet,
      rows: [
        ...sheet.rows,
        {
          id: newRowId,
          excelRowNumber: (sheet.rows[sheet.rows.length - 1]?.excelRowNumber || sheet.rows.length) + 1,
          cells: initialCells,
        },
      ],
    }));

    if (activeSheet.fields[0]) {
      setSelectedCell({ rowId: newRowId, fieldId: activeSheet.fields[0].id });
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
    const label = isHeaderless ? colLetter : `Field ${activeSheet.fields.length + 1}`;
    const newField: FieldConfig = {
      id: `fld_${Date.now().toString(36)}`,
      key: isHeaderless ? `col_${colLetter.toLowerCase()}` : `field_${activeSheet.fields.length + 1}`,
      label,
      sourceHeader: isHeaderless ? `(No Header — Col ${colLetter})` : label,
      columnLetter: colLetter,
      type: 'text',
      visible: true,
      width: 150,
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
          width: 160,
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
          width: 160,
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
          width: 160,
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
    <div className="app-shell">
      {/* Top Editorial Bar */}
      <header className="top-navbar">
        <div className="brand-cluster">
          <div className="os-dots" aria-hidden="true">
            <span className="os-dot" />
            <span className="os-dot" />
            <span className="os-dot" />
          </div>
          <span className="brand-title">Tabula Loom</span>
          <div className="divider-vertical" />
          <input
            type="text"
            className="workbook-title-input"
            value={project.fileName}
            onChange={(e) => setProject((p) => ({ ...p, fileName: e.target.value }))}
            aria-label="Workbook file name"
          />
          {totalErrorCount === 0 ? (
            <span className="status-pill valid">
              <IconCheck size={11} /> Schema Valid
            </span>
          ) : (
            <span
              className="status-pill invalid"
              title="Click to inspect validation issues"
              onClick={() => {
                const firstErr = Object.values(activeSheetErrors)[0];
                if (firstErr) {
                  setSelectedCell({ rowId: firstErr.rowId, fieldId: firstErr.fieldId });
                  setSelectedFieldId(firstErr.fieldId);
                }
              }}
            >
              <IconWarning size={11} /> {totalErrorCount} Validation{' '}
              {totalErrorCount === 1 ? 'Issue' : 'Issues'}
            </span>
          )}
          {saveToastMessage && (
            <span className="status-pill info" style={{ textTransform: 'none' }}>
              {saveToastMessage}
            </span>
          )}
        </div>

        {/* Center Mode Switcher */}
        <div className="mode-switcher" role="tablist" aria-label="Workspace view mode">
          <button
            type="button"
            role="tab"
            aria-selected={workspaceMode === 'grid'}
            className={`mode-btn ${workspaceMode === 'grid' ? 'active' : ''}`}
            onClick={() => setWorkspaceMode('grid')}
          >
            <IconTable size={14} />
            Spreadsheet & Inspector
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={workspaceMode === 'schema'}
            className={`mode-btn ${workspaceMode === 'schema' ? 'active' : ''}`}
            onClick={() => setWorkspaceMode('schema')}
          >
            <IconSliders size={14} />
            Fields Configuration ({activeSheet.fields.length})
          </button>
        </div>

        {/* Right File & PostgreSQL Actions */}
        <div className="top-actions">
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
            onClick={() => {
              fetchBackendStatusAndWorkbooks();
              dbDrawerDialogRef.current?.showModal();
            }}
            title="Manage saved workbooks in PostgreSQL & Railway Storage"
          >
            <IconDatabase size={14} />
            PostgreSQL ({savedWorkbooks.length})
            <span
              style={{
                width: '7px',
                height: '7px',
                borderRadius: '50%',
                backgroundColor: dbStatus.connected ? '#346538' : '#956400',
              }}
            />
          </button>
          <button
            type="button"
            className="btn"
            disabled={isSavingToDb}
            onClick={handleSaveToPostgres}
            title="Save all sheets, dynamic field schemas, and rows to PostgreSQL"
          >
            {isSavingToDb ? 'Saving...' : project.id ? 'Sync to DB' : 'Save to DB'}
          </button>
          <button
            type="button"
            className="btn"
            onClick={() => fileInputRef.current?.click()}
            title="Import an .xlsx, .xls, or .csv workbook"
          >
            <IconUpload size={14} />
            Import .xlsx
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={() => exportDialogRef.current?.showModal()}
          >
            <IconDownload size={14} />
            Export .xlsx
          </button>
        </div>
      </header>

      {/* Hidden Sheets Detection Notice Banner */}
      {hiddenSheets.length > 0 && !dismissedHiddenNotice && (
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '7px 20px',
            backgroundColor: 'var(--pastel-yellow-bg)',
            borderBottom: '1px solid #EFE0B4',
            color: 'var(--pastel-yellow-fg)',
            fontSize: '12.5px',
            flexShrink: 0,
            gap: '12px',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
            <IconWarning size={14} />
            <span>
              <strong>
                {hiddenSheets.length} hidden{' '}
                {hiddenSheets.length === 1 ? 'worksheet was' : 'worksheets were'} found
              </strong>{' '}
              in <code>{project.fileName}</code>:{' '}
              {hiddenSheets.map((hs, idx) => (
                <React.Fragment key={hs.id}>
                  <code
                    style={{
                      background: 'rgba(255,255,255,0.65)',
                      padding: '1px 6px',
                      borderRadius: '4px',
                      fontWeight: 600,
                    }}
                  >
                    {hs.name}
                    {hs.hiddenState === 2 ? ' (VeryHidden)' : ' (Hidden)'}
                  </code>
                  {idx < hiddenSheets.length - 1 ? ', ' : ''}
                </React.Fragment>
              ))}
            </span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
            <button
              type="button"
              className="btn btn-sm"
              style={{
                backgroundColor: '#FFFFFF',
                borderColor: '#DFC98B',
                color: '#785000',
              }}
              onClick={() => setShowHiddenSheets((prev) => !prev)}
            >
              <IconEye visible={showHiddenSheets} size={13} />
              {showHiddenSheets
                ? 'Hide Hidden Sheets'
                : `Show Hidden ${hiddenSheets.length === 1 ? 'Sheet' : 'Sheets'} (${hiddenSheets.length})`}
            </button>
            <button
              type="button"
              className="btn btn-ghost btn-sm"
              style={{ color: '#785000' }}
              onClick={() => setDismissedHiddenNotice(true)}
            >
              Dismiss
            </button>
          </div>
        </div>
      )}

      {/* Secondary Sheet Navigation & Quick Toolbar */}
      <div className="sheet-toolbar">
        <div className="sheet-tabs-bar" role="tablist" aria-label="Workbook sheets">
          {displayedSheets.map((sheet) => {
            const sheetErrCount = Object.keys(sheetErrorsMap[sheet.id] || {}).length;
            const isActive = sheet.id === activeSheet.id;
            const isHiddenSheet = Boolean(sheet.hiddenState);
            const isNoHeader = sheet.headerRowIndex === -1;
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
                        borderColor: isActive ? '#956400' : 'var(--border-strong)',
                        opacity: isActive ? 1 : 0.78,
                      }
                    : undefined
                }
              >
                {isHiddenSheet && (
                  <span
                    className="status-pill warn"
                    style={{ padding: '0px 5px', fontSize: '9.5px' }}
                    title="This worksheet is marked Hidden in the Excel file"
                  >
                    HIDDEN
                  </span>
                )}
                <span>{sheet.name}</span>
                {isNoHeader && (
                  <span
                    style={{
                      fontFamily: 'var(--font-mono)',
                      fontSize: '10px',
                      color: 'var(--text-muted)',
                      border: '1px solid var(--border-hairline)',
                      borderRadius: '3px',
                      padding: '0 4px',
                    }}
                    title="No header row detected — using column letters A, B, C..."
                  >
                    A·B·C
                  </span>
                )}
                <span
                  className="sheet-tab-count"
                  title={`${sheet.fields.length} columns, ${sheet.rows.length} rows`}
                >
                  {sheet.fields.length}c · {sheet.rows.length}r
                </span>
                {sheetErrCount > 0 && (
                  <span
                    style={{
                      color: 'var(--pastel-red-fg)',
                      fontSize: '11px',
                      fontFamily: 'var(--font-mono)',
                      fontWeight: 600,
                    }}
                  >
                    !{sheetErrCount}
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
              title={`Hidden sheets: ${hiddenSheets.map((s) => s.name).join(', ')}`}
            >
              <IconEye visible={showHiddenSheets} size={13} />
              {showHiddenSheets
                ? `Showing ${hiddenSheets.length} Hidden`
                : `+${hiddenSheets.length} Hidden`}
            </button>
          )}
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => newSheetDialogRef.current?.showModal()}
            title="Create a new worksheet"
          >
            <IconPlus size={13} /> New Sheet
          </button>
        </div>

        <div className="toolbar-right">
          {/* Dynamic Header Mode Selector */}
          <div
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              fontSize: '12px',
              color: 'var(--text-muted)',
              background: 'var(--bg-surface)',
              border: '1px solid var(--border-hairline)',
              borderRadius: '5px',
              padding: '0 8px',
              height: '28px',
            }}
            title="Choose whether this sheet has a header row or should use Excel column letters (A, B, C...)"
          >
            <span>Headers:</span>
            <select
              value={activeSheet.headerRowIndex}
              onChange={(e) => handleHeaderRowChange(Number(e.target.value))}
              style={{
                border: 'none',
                background: 'transparent',
                fontFamily: 'var(--font-mono)',
                fontSize: '12px',
                fontWeight: 600,
                color: 'var(--text-primary)',
                cursor: 'pointer',
              }}
            >
              <option value={-1}>None — Use A, B, C...</option>
              {(activeSheet.rawMatrix || []).slice(0, 12).map((r, idx) => {
                const preview = r
                  .filter((c) => String(c ?? '').trim() !== '')
                  .slice(0, 3)
                  .join(', ');
                return (
                  <option key={idx} value={idx}>
                    Row {idx + 1} {preview ? `(${preview.slice(0, 24)})` : '(empty)'}
                  </option>
                );
              })}
              {!activeSheet.rawMatrix && <option value={0}>Row 1 (Default)</option>}
            </select>
          </div>

          <div className="search-box">
            <IconSearch size={13} />
            <input
              type="text"
              placeholder={`Filter ${activeSheet.name}...`}
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
            />
          </div>
          <button
            type="button"
            className="btn btn-sm"
            onClick={() => setShowHiddenColumns((prev) => !prev)}
            title="Toggle visibility of hidden columns in the grid"
          >
            <IconEye visible={showHiddenColumns} size={13} />
            {showHiddenColumns ? 'Showing Hidden Cols' : 'Hide Disabled Cols'}
          </button>
          <button type="button" className="btn btn-sm" onClick={handleAddField}>
            <IconPlus size={13} /> Add Col
          </button>
          <button type="button" className="btn btn-sm btn-primary" onClick={handleAddRow}>
            <IconPlus size={13} /> Add Row
          </button>
        </div>
      </div>

      {/* Main Workspace Body */}
      <div className="workspace-body">
        {workspaceMode === 'grid' ? (
          <>
            {/* Left / Center Data Grid */}
            <main className="primary-stage">
              {/* Active Cell / Formula Bar */}
              <div className="formula-bar">
                <span className="cell-coord-badge">
                  {activeCellInfo ? activeCellInfo.coord : '—'}
                </span>
                {activeCellInfo ? (
                  <>
                    <FieldTypeGlyph type={activeCellInfo.field.type} />
                    <span
                      style={{
                        fontFamily: 'var(--font-mono)',
                        fontSize: '11.5px',
                        color: 'var(--text-muted)',
                      }}
                    >
                      {activeCellInfo.field.key}
                    </span>
                    <div className="divider-vertical" />
                    {activeCellInfo.field.type === 'formula' ? (
                      <input
                        type="text"
                        className="formula-bar-input"
                        value={activeCellInfo.field.formulaExpression || ''}
                        placeholder="Enter column formula, e.g. {unit_cost} * {stock_qty}"
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
                        placeholder={`Edit value for ${activeCellInfo.field.label}...`}
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
                      <span className="status-pill invalid" style={{ textTransform: 'none' }}>
                        <IconWarning size={11} /> {activeCellInfo.err.message}
                      </span>
                    )}
                    {!isHeaderlessSheet && (
                      <button
                        type="button"
                        className="btn btn-ghost btn-sm"
                        onClick={() => handleHeaderRowChange(-1)}
                        title="Treat this sheet as having no headers (use A, B, C... and restore header row as data)"
                      >
                        Clear Header (Use A, B, C)
                      </button>
                    )}
                    <button
                      type="button"
                      className="btn btn-ghost btn-sm"
                      onClick={() => handlePromoteDataRowToHeader(activeCellInfo.row.id)}
                      title="Use this row's values as the column headers for this sheet"
                    >
                      Set Row {activeCellInfo.displayRowNum} as Header
                    </button>
                  </>
                ) : (
                  <span className="formula-hint">Select any cell or column header to inspect</span>
                )}
              </div>

              {/* Scrollable Multi-Sheet Grid */}
              <div className="grid-scroll-container">
                <table className="sheet-table">
                  <thead>
                    <tr>
                      <th className="th-row-index">#</th>
                      {visibleFields.map((field, colIdx) => {
                        const isSelectedCol = field.id === selectedFieldId;
                        const colLetter = field.columnLetter || getExcelColumnLetter(colIdx);
                        const isJustAlphabet = field.label === colLetter;
                        return (
                          <th
                            key={field.id}
                            style={{ width: field.width, minWidth: field.width }}
                            className={isSelectedCol ? 'selected-column' : ''}
                            onClick={() => setSelectedFieldId(field.id)}
                          >
                            <div className="th-header-content">
                              <div className="th-top-line">
                                <div className="th-label-group">
                                  {!isJustAlphabet && (
                                    <span
                                      style={{
                                        fontFamily: 'var(--font-mono)',
                                        fontSize: '10px',
                                        color: 'var(--text-muted)',
                                      }}
                                    >
                                      {colLetter}
                                    </span>
                                  )}
                                  <FieldTypeGlyph type={field.type} />
                                  <span
                                    className="th-label"
                                    style={
                                      isJustAlphabet
                                        ? { fontFamily: 'var(--font-mono)', fontSize: '13px' }
                                        : undefined
                                    }
                                  >
                                    {field.label}
                                  </span>
                                  {field.validation.required && (
                                    <span className="req-dot" title="Required field">
                                      *
                                    </span>
                                  )}
                                </div>
                                {!field.visible && (
                                  <span
                                    style={{
                                      fontSize: '10px',
                                      color: 'var(--text-muted)',
                                      fontFamily: 'var(--font-mono)',
                                    }}
                                  >
                                    HIDDEN
                                  </span>
                                )}
                              </div>
                              <div className="th-key-subtitle">
                                {field.key}
                                {field.type === 'lookup' && field.lookup && (
                                  <span>
                                    {' '}
                                    →{' '}
                                    {project.sheets.find((s) => s.id === field.lookup?.targetSheetId)
                                      ?.name || 'Sheet'}
                                  </span>
                                )}
                                {field.type === 'formula' && field.formulaExpression && (
                                  <span> = {field.formulaExpression}</span>
                                )}
                              </div>
                            </div>
                          </th>
                        );
                      })}
                    </tr>
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
                                title="Delete row"
                                onClick={() => handleDeleteRow(row.id)}
                              >
                                <IconTrash size={12} />
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
                                      <option value="">— Select Foreign Key —</option>
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
                                      <option value="true">True</option>
                                      <option value="false">False</option>
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
                                        : 'Double-click to edit cell value'
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
                                          <IconLink size={11} />
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
                                        style={{ cursor: 'pointer', accentColor: '#111111' }}
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
                                        <IconWarning size={13} />
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

            {/* Right Field Configuration Inspector Sidebar */}
            {selectedField && (
              <aside className="field-inspector" aria-label="Field Configuration Inspector">
                <div className="inspector-header">
                  <div>
                    <div className="section-heading">
                      Field Configuration • Col {selectedField.columnLetter || '—'}
                    </div>
                    <h2 className="inspector-title">
                      {selectedField.label === selectedField.columnLetter
                        ? `Column ${selectedField.columnLetter}`
                        : selectedField.label}
                    </h2>
                  </div>
                  <div style={{ display: 'flex', gap: '4px' }}>
                    <button
                      type="button"
                      className="btn btn-sm"
                      onClick={() => handleMoveField(selectedField.id, -1)}
                      title="Move column left"
                    >
                      ←
                    </button>
                    <button
                      type="button"
                      className="btn btn-sm"
                      onClick={() => handleMoveField(selectedField.id, 1)}
                      title="Move column right"
                    >
                      →
                    </button>
                  </div>
                </div>

                {/* Section 1: Mapping & Identity */}
                <div className="inspector-section">
                  <div className="section-heading">Identity & Header Mapping</div>

                  <div className="form-field">
                    <label className="form-label">Display Label (Column Header)</label>
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
                    <label className="form-label">
                      <span>Machine / Schema Key</span>
                      <button
                        type="button"
                        className="btn btn-ghost btn-sm"
                        style={{ height: '20px', padding: '0 6px', fontSize: '11px' }}
                        onClick={() =>
                          updateFieldConfig(selectedField.id, {
                            key: toMachineKey(selectedField.label),
                          })
                        }
                      >
                        Auto-slug
                      </button>
                    </label>
                    <input
                      type="text"
                      className="input-control mono"
                      value={selectedField.key}
                      onChange={(e) =>
                        updateFieldConfig(selectedField.id, { key: e.target.value })
                      }
                    />
                    <span className="form-hint">
                      {isHeaderlessSheet ? (
                        <>
                          Sheet has no header row — using Excel column{' '}
                          <code>{selectedField.columnLetter}</code>
                        </>
                      ) : (
                        <>
                          Original .xlsx cell:{' '}
                          <code>
                            {selectedField.columnLetter || ''}
                            {activeSheet.headerRowIndex + 1} = {selectedField.sourceHeader}
                          </code>
                        </>
                      )}
                    </span>
                  </div>

                  <div className="form-row-2">
                    <label className="checkbox-row">
                      <span>Visible in Grid</span>
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
                        min={90}
                        max={420}
                        value={selectedField.width}
                        onChange={(e) =>
                          updateFieldConfig(selectedField.id, {
                            width: Number(e.target.value) || 140,
                          })
                        }
                        title="Column width in pixels"
                      />
                    </div>
                  </div>
                </div>

                {/* Section 2: Data Type & Type-Specific Rules */}
                <div className="inspector-section">
                  <div className="section-heading">Data Type & Formatting</div>

                  <div className="form-field">
                    <label className="form-label">Data Type</label>
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
                      <option value="text">Text (String)</option>
                      <option value="number">Number (Integer / Decimal / Currency)</option>
                      <option value="date">Date (YYYY-MM-DD)</option>
                      <option value="select">Select / Enum (Allowed Options)</option>
                      <option value="boolean">Boolean (True / False)</option>
                      <option value="formula">Computed Formula (fx)</option>
                      <option value="lookup">Cross-Sheet Lookup (Foreign Key)</option>
                    </select>
                  </div>

                  {(selectedField.type === 'number' || selectedField.type === 'formula') && (
                    <div className="form-row-2">
                      <div className="form-field">
                        <label className="form-label">Number Format</label>
                        <select
                          className="select-control"
                          value={selectedField.numberFormat || 'raw'}
                          onChange={(e) =>
                            updateFieldConfig(selectedField.id, {
                              numberFormat: e.target.value as NumberFormatType,
                            })
                          }
                        >
                          <option value="raw">Raw Number</option>
                          <option value="integer">Integer (1,000)</option>
                          <option value="decimal">Decimal (1,000.00)</option>
                          <option value="currency">Currency ($1,000.00)</option>
                          <option value="percent">Percentage (10.0%)</option>
                        </select>
                      </div>
                      <div className="form-field">
                        <label className="form-label">Currency Symbol</label>
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
                      <label className="form-label">Allowed Options (comma-separated)</label>
                      <textarea
                        className="textarea-control mono"
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
                      <span className="form-hint">
                        Cells with values outside this list will be flagged.
                      </span>
                    </div>
                  )}

                  {selectedField.type === 'formula' && (
                    <div className="form-field">
                      <label className="form-label">Formula Expression</label>
                      <input
                        type="text"
                        className="input-control mono"
                        placeholder="{unit_cost} * {stock_qty}"
                        value={selectedField.formulaExpression || ''}
                        onChange={(e) =>
                          updateFieldConfig(selectedField.id, {
                            formulaExpression: e.target.value,
                          })
                        }
                      />
                      <span className="form-hint">
                        Reference sibling columns by machine key in curly braces:{' '}
                        {activeSheet.fields
                          .filter((f) => f.id !== selectedField.id)
                          .slice(0, 4)
                          .map((f) => `{${f.key}}`)
                          .join(', ')}
                      </span>
                    </div>
                  )}

                  {selectedField.type === 'lookup' && (
                    <>
                      <div className="form-field">
                        <label className="form-label">Target Sheet</label>
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
                              <label className="form-label">Key Column</label>
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
                                    {tf.label} ({tf.key})
                                  </option>
                                ))}
                              </select>
                            </div>
                            <div className="form-field">
                              <label className="form-label">Preview Label</label>
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

                  <div className="form-field">
                    <label className="form-label">Default Value for New Rows</label>
                    <input
                      type="text"
                      className="input-control mono"
                      placeholder="Empty"
                      value={selectedField.defaultValue}
                      onChange={(e) =>
                        updateFieldConfig(selectedField.id, { defaultValue: e.target.value })
                      }
                    />
                  </div>
                </div>

                {/* Section 3: Validation Constraints */}
                <div className="inspector-section">
                  <div className="section-heading">Validation Rules</div>

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
                      <span>Unique</span>
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
                        <label className="form-label">Min Value</label>
                        <input
                          type="number"
                          className="input-control mono"
                          placeholder="None"
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
                        <label className="form-label">Max Value</label>
                        <input
                          type="number"
                          className="input-control mono"
                          placeholder="None"
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

                  {selectedField.type === 'text' && (
                    <div className="form-field">
                      <label className="form-label">Regex Pattern Match</label>
                      <input
                        type="text"
                        className="input-control mono"
                        placeholder="e.g. ^[A-Z]{3}-\d{4}$"
                        value={selectedField.validation.regex || ''}
                        onChange={(e) =>
                          updateFieldConfig(selectedField.id, {
                            validation: {
                              ...selectedField.validation,
                              regex: e.target.value || undefined,
                            },
                          })
                        }
                      />
                    </div>
                  )}
                </div>

                {/* Section 4: Notes & Delete */}
                <div className="inspector-section" style={{ borderBottom: 'none' }}>
                  <div className="form-field">
                    <label className="form-label">Field Documentation Note</label>
                    <input
                      type="text"
                      className="input-control"
                      placeholder="Describe business rules for this column..."
                      value={selectedField.description || ''}
                      onChange={(e) =>
                        updateFieldConfig(selectedField.id, { description: e.target.value })
                      }
                    />
                  </div>

                  <button
                    type="button"
                    className="btn btn-danger"
                    style={{ marginTop: '6px' }}
                    disabled={activeSheet.fields.length <= 1}
                    onClick={() => handleDeleteField(selectedField.id)}
                  >
                    <IconTrash size={13} /> Delete Column from Sheet
                  </button>
                </div>
              </aside>
            )}
          </>
        ) : (
          /* Dedicated Full-Screen Fields Configuration Matrix View */
          <main className="schema-view-container">
            <div className="schema-view-inner">
              <div className="schema-hero-card">
                <div>
                  <div className="section-heading">
                    Worksheet Schema & Mapping •{' '}
                    {isHeaderlessSheet
                      ? 'No Header Row (Columns A, B, C...)'
                      : `Header Row ${activeSheet.headerRowIndex + 1}`}
                  </div>
                  <h1 className="schema-hero-title">
                    {activeSheet.name}{' '}
                    {activeSheet.hiddenState ? (
                      <span className="status-pill warn" style={{ verticalAlign: 'middle' }}>
                        Hidden Sheet
                      </span>
                    ) : null}
                  </h1>
                  <p style={{ color: 'var(--text-muted)', fontSize: '13.5px' }}>
                    {activeSheet.description ||
                      'Configure field data types, machine export keys, cross-sheet foreign keys, and validation constraints.'}
                  </p>
                </div>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <button
                    type="button"
                    className="btn"
                    onClick={() => exportSchemaAsJson(project)}
                  >
                    Export Schema JSON
                  </button>
                  <button type="button" className="btn" onClick={handleAddField}>
                    <IconPlus size={14} /> Add Field Column
                  </button>
                  <button
                    type="button"
                    className="btn btn-primary"
                    onClick={() => setWorkspaceMode('grid')}
                  >
                    <IconTable size={14} /> Return to Data Grid
                  </button>
                </div>
              </div>

              {/* Bento Summary Metrics */}
              <div className="schema-bento-stats">
                <div className="stat-card">
                  <div className="stat-label">Configured Fields</div>
                  <div className="stat-value">{activeSheet.fields.length}</div>
                </div>
                <div className="stat-card">
                  <div className="stat-label">Header Mode</div>
                  <div className="stat-value" style={{ fontSize: '17px', paddingTop: '4px' }}>
                    {isHeaderlessSheet ? 'Alphabet (A, B, C)' : `Row ${activeSheet.headerRowIndex + 1}`}
                  </div>
                </div>
                <div className="stat-card">
                  <div className="stat-label">Required Constraints</div>
                  <div className="stat-value">
                    {activeSheet.fields.filter((f) => f.validation.required).length}
                  </div>
                </div>
                <div className="stat-card">
                  <div className="stat-label">Validation Health</div>
                  <div className="stat-value">
                    {Object.keys(activeSheetErrors).length === 0
                      ? '100%'
                      : `${Object.keys(activeSheetErrors).length} errs`}
                  </div>
                </div>
              </div>

              {/* Full Schema Table */}
              <div className="schema-table-card">
                <table className="schema-config-table">
                  <thead>
                    <tr>
                      <th style={{ width: '48px' }}>Vis</th>
                      <th style={{ width: '52px' }}>Col</th>
                      <th style={{ width: '190px' }}>Display Label</th>
                      <th style={{ width: '170px' }}>Machine Key</th>
                      <th style={{ width: '150px' }}>Data Type</th>
                      <th>Type Parameters / Lookup / Formula</th>
                      <th style={{ width: '125px' }}>Default</th>
                      <th style={{ width: '140px' }}>Rules</th>
                      <th style={{ width: '60px' }} />
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
                            title={field.visible ? 'Visible in grid' : 'Hidden in grid'}
                          >
                            <IconEye visible={field.visible} size={14} />
                          </button>
                        </td>
                        <td
                          style={{
                            fontFamily: 'var(--font-mono)',
                            fontSize: '12px',
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
                              updateFieldConfig(field.id, { label: e.target.value })
                            }
                          />
                        </td>
                        <td>
                          <input
                            type="text"
                            className="input-control mono"
                            value={field.key}
                            onChange={(e) =>
                              updateFieldConfig(field.id, { key: e.target.value })
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
                            <option value="number">Number</option>
                            <option value="date">Date</option>
                            <option value="select">Select / Enum</option>
                            <option value="boolean">Boolean</option>
                            <option value="formula">Formula (fx)</option>
                            <option value="lookup">Lookup (Ref)</option>
                          </select>
                        </td>
                        <td>
                          {field.type === 'select' && (
                            <input
                              type="text"
                              className="input-control mono"
                              placeholder="Opt1, Opt2, Opt3"
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
                              <option value="raw">Raw Number</option>
                              <option value="integer">Integer (1,000)</option>
                              <option value="decimal">Decimal (1,000.00)</option>
                              <option value="currency">Currency ($1,000.00)</option>
                              <option value="percent">Percent (10.0%)</option>
                            </select>
                          )}
                          {field.type === 'lookup' && (
                            <div style={{ display: 'flex', gap: '6px' }}>
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
                                    Key: {tf.key}
                                  </option>
                                ))}
                              </select>
                            </div>
                          )}
                          {(field.type === 'text' ||
                            field.type === 'date' ||
                            field.type === 'boolean') && (
                            <span className="form-hint">
                              Source: <code>{field.sourceHeader}</code>
                            </span>
                          )}
                        </td>
                        <td>
                          <input
                            type="text"
                            className="input-control mono"
                            placeholder="None"
                            value={field.defaultValue}
                            onChange={(e) =>
                              updateFieldConfig(field.id, {
                                defaultValue: e.target.value,
                              })
                            }
                          />
                        </td>
                        <td>
                          <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                            <label
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '4px',
                                fontSize: '12px',
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
                              Req
                            </label>
                            <label
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '4px',
                                fontSize: '12px',
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
                              Uniq
                            </label>
                          </div>
                        </td>
                        <td>
                          <button
                            type="button"
                            className="btn btn-ghost btn-sm"
                            disabled={activeSheet.fields.length <= 1}
                            onClick={() => handleDeleteField(field.id)}
                            title="Remove field"
                          >
                            <IconTrash size={13} />
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

      {/* Bottom Status Bar */}
      <footer className="bottom-statusbar">
        <div className="statusbar-group">
          <span>
            Sheet: <strong>{activeSheet.name}</strong>
            {activeSheet.hiddenState ? ' (Hidden in Excel)' : ''} ({filteredRows.length} of{' '}
            {activeSheet.rows.length} rows •{' '}
            {isHeaderlessSheet
              ? 'No Header Row [A, B, C...]'
              : `Header at Excel Row ${activeSheet.headerRowIndex + 1}`}
            )
          </span>
          <span>•</span>
          <span>
            Visible Columns: {visibleFields.length} / {activeSheet.fields.length}
          </span>
          <span>•</span>
          <span>
            DB: {dbStatus.connected ? 'PostgreSQL Connected' : 'Local Mode (Set DATABASE_URL)'}
          </span>
        </div>
        <div className="statusbar-group">
          <span>
            Double-click cell to edit • Press <kbd className="keycap">Enter</kbd> to commit •{' '}
            <kbd className="keycap">Esc</kbd> to cancel
          </span>
        </div>
      </footer>

      {/* Modal Dialog: PostgreSQL & Railway Storage Manager */}
      <dialog
        ref={dbDrawerDialogRef}
        closedby="any"
        className="editorial-dialog"
        style={{ maxWidth: '640px' }}
      >
        <div className="dialog-header">
          <div>
            <div className="section-heading">Database & Storage</div>
            <h3 className="dialog-title">PostgreSQL Workbooks & Dynamic Schemas</h3>
          </div>
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
              padding: '12px 14px',
              borderRadius: '6px',
              border: '1px solid var(--border-hairline)',
              backgroundColor: 'var(--bg-subtle)',
            }}
          >
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span
                  className={`status-pill ${dbStatus.connected ? 'valid' : 'warn'}`}
                >
                  {dbStatus.connected ? 'PostgreSQL Connected' : 'PostgreSQL Not Connected'}
                </span>
                <span
                  style={{
                    fontFamily: 'var(--font-mono)',
                    fontSize: '11.5px',
                    color: 'var(--text-muted)',
                  }}
                >
                  Storage: {dbStatus.storageMode}
                </span>
              </div>
              {dbStatus.error && (
                <div style={{ fontSize: '12px', color: 'var(--text-muted)', marginTop: '4px' }}>
                  {dbStatus.error}
                </div>
              )}
            </div>
            <button
              type="button"
              className="btn btn-primary btn-sm"
              disabled={!dbStatus.connected || isSavingToDb}
              onClick={handleSaveToPostgres}
            >
              {isSavingToDb
                ? 'Saving...'
                : project.id
                ? 'Update Current Workbook'
                : 'Save Current Workbook'}
            </button>
          </div>

          <div className="section-heading" style={{ marginTop: '4px' }}>
            Saved Workbooks in PostgreSQL ({savedWorkbooks.length})
          </div>

          {savedWorkbooks.length === 0 ? (
            <div
              style={{
                padding: '24px',
                textAlign: 'center',
                border: '1px dashed var(--border-strong)',
                borderRadius: '6px',
                color: 'var(--text-muted)',
                fontSize: '13px',
              }}
            >
              {dbStatus.connected
                ? 'No workbooks saved yet. Click "Save Current Workbook" to store all sheets, dynamic field configs, and rows in PostgreSQL.'
                : 'Add DATABASE_URL to your .env file (or link a Railway PostgreSQL database) and restart yarn dev.'}
            </div>
          ) : (
            <div
              style={{
                display: 'flex',
                flexDirection: 'column',
                gap: '8px',
                maxHeight: '280px',
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
                      padding: '10px 14px',
                      border: '1px solid var(--border-hairline)',
                      borderRadius: '6px',
                      backgroundColor: isCurrent ? 'var(--bg-selected)' : 'var(--bg-surface)',
                    }}
                  >
                    <div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <strong style={{ fontSize: '13px', color: 'var(--text-primary)' }}>
                          {wb.fileName}
                        </strong>
                        <span className="sheet-tab-count">{wb.sheetCount} sheets</span>
                        {isCurrent && <span className="status-pill info">Active</span>}
                      </div>
                      <div
                        style={{
                          fontFamily: 'var(--font-mono)',
                          fontSize: '11px',
                          color: 'var(--text-muted)',
                          marginTop: '2px',
                        }}
                      >
                        ID: {wb.id} • Updated {new Date(wb.lastModified).toLocaleString()}
                      </div>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      {wb.storageKey && (
                        <a
                          href={`/api/workbooks/${wb.id}/original-file`}
                          className="btn btn-sm"
                          style={{ textDecoration: 'none' }}
                          title="Download original uploaded .xlsx binary from storage"
                        >
                          Original .xlsx
                        </a>
                      )}
                      <button
                        type="button"
                        className="btn btn-sm"
                        disabled={isLoadingFromDb}
                        onClick={() => handleLoadWorkbookFromDb(wb.id)}
                      >
                        Load
                      </button>
                      <button
                        type="button"
                        className="btn btn-ghost btn-sm"
                        onClick={() => handleDeleteWorkbookFromDb(wb.id, wb.fileName)}
                        title="Delete from PostgreSQL"
                      >
                        <IconTrash size={13} />
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
            className="btn"
            onClick={() => exportSchemaAsJson(project)}
          >
            Download Schema JSON
          </button>
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
          <h3 className="dialog-title">Export Configured Workbook (.xlsx)</h3>
          <button
            type="button"
            className="btn btn-ghost btn-sm"
            onClick={() => exportDialogRef.current?.close()}
          >
            Close
          </button>
        </div>
        <div className="dialog-body">
          <p style={{ fontSize: '13px', color: 'var(--text-muted)' }}>
            Choose how field mappings, computed formulas, and schema configurations are written to
            the exported Excel file.
          </p>

          <label className="checkbox-row">
            <div>
              <span>Use Machine Keys as Column Headers</span>
              <div className="form-hint">
                Exports snake_case keys (e.g. <code>unit_cost</code>) instead of display labels
              </div>
            </div>
            <input
              type="checkbox"
              checked={exportUseMachineKeys}
              onChange={(e) => setExportUseMachineKeys(e.target.checked)}
            />
          </label>

          <label className="checkbox-row">
            <div>
              <span>Include Hidden Columns</span>
              <div className="form-hint">
                Export all {activeSheet.fields.length} columns even if hidden in the grid view
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
                <span>Include Hidden Excel Sheets ({hiddenSheets.length})</span>
                <div className="form-hint">
                  Exports hidden worksheets ({hiddenSheets.map((s) => s.name).join(', ')}) preserving
                  their Hidden state
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
              <span>Append &ldquo;_field_configs&rdquo; Metadata Sheet</span>
              <div className="form-hint">
                Includes a dedicated worksheet documenting all field types, lookups, and rules
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
            <IconDownload size={14} /> Download .xlsx
          </button>
        </div>
      </dialog>

      {/* Modal Dialog: Create New Sheet */}
      <dialog ref={newSheetDialogRef} closedby="any" className="editorial-dialog">
        <form onSubmit={handleCreateSheet}>
          <div className="dialog-header">
            <h3 className="dialog-title">Create New Worksheet</h3>
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
              <label className="form-label">Worksheet Name</label>
              <input
                type="text"
                className="input-control"
                placeholder="e.g. Warehouse Locations"
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
    </div>
  );
}
