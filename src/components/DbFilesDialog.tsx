import type { WorkbookProject, SavedWorkbookSummary } from '../types';
import type { DbStatus } from '../types';
import { IconTrash } from '../Icons';

interface DbFilesDialogProps {
  dbDrawerDialogRef: React.RefObject<HTMLDialogElement | null>;
  dbStatus: DbStatus;
  handleDeleteWorkbookFromDb: (workbookId: string, fileName: string) => void | Promise<void>;
  handleLoadWorkbookFromDb: (workbookId: string) => void | Promise<void>;
  isLoadingFromDb: boolean;
  project: WorkbookProject;
  savedWorkbooks: SavedWorkbookSummary[];
}

export function DbFilesDialog({ dbDrawerDialogRef, dbStatus, handleDeleteWorkbookFromDb, handleLoadWorkbookFromDb, isLoadingFromDb, project, savedWorkbooks }: DbFilesDialogProps) {
  return (
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
  );
}
