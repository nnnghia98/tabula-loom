import type { WorkbookProject, SavedWorkbookSummary, SheetData } from '../types';
import type { WorkspaceMode } from '../types';
import type { SheetErrors } from '../types';
import { IconUpload, IconDownload, IconSliders, IconTable, IconWarning, IconDatabase } from '../Icons';

interface TopNavbarProps {
  activeSheet: SheetData;
  activeSheetErrors: SheetErrors;
  dbDrawerDialogRef: React.RefObject<HTMLDialogElement | null>;
  draftState: 'idle' | 'saved';
  exportDialogRef: React.RefObject<HTMLDialogElement | null>;
  fetchBackendStatusAndWorkbooks: () => Promise<void>;
  fileInputRef: React.RefObject<HTMLInputElement | null>;
  handleFileUpload: (e: React.ChangeEvent<HTMLInputElement>) => void | Promise<void>;
  handleSaveToPostgres: () => void | Promise<void>;
  isSavingToDb: boolean;
  project: WorkbookProject;
  saveToastMessage: string | null;
  savedWorkbooks: SavedWorkbookSummary[];
  setProject: React.Dispatch<React.SetStateAction<WorkbookProject>>;
  setSelectedCell: React.Dispatch<React.SetStateAction<{ rowId: string; fieldId: string } | null>>;
  setSelectedFieldId: React.Dispatch<React.SetStateAction<string>>;
  setUiScale: React.Dispatch<React.SetStateAction<number>>;
  setWorkspaceMode: React.Dispatch<React.SetStateAction<WorkspaceMode>>;
  totalErrorCount: number;
  uiScale: number;
  workspaceMode: WorkspaceMode;
}

export function TopNavbar({ activeSheet, activeSheetErrors, dbDrawerDialogRef, draftState, exportDialogRef, fetchBackendStatusAndWorkbooks, fileInputRef, handleFileUpload, handleSaveToPostgres, isSavingToDb, project, saveToastMessage, savedWorkbooks, setProject, setSelectedCell, setSelectedFieldId, setUiScale, setWorkspaceMode, totalErrorCount, uiScale, workspaceMode }: TopNavbarProps) {
  return (
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
  );
}
