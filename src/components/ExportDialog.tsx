import type { WorkbookProject, SheetData } from '../types';
import { exportWorkbookToXlsx } from '../xlsxUtils';
import { IconDownload } from '../Icons';

interface ExportDialogProps {
  activeSheet: SheetData;
  exportDialogRef: React.RefObject<HTMLDialogElement | null>;
  exportIncludeHidden: boolean;
  exportIncludeHiddenSheets: boolean;
  exportIncludeSchemaTab: boolean;
  exportUseMachineKeys: boolean;
  hiddenSheets: SheetData[];
  project: WorkbookProject;
  setExportIncludeHidden: React.Dispatch<React.SetStateAction<boolean>>;
  setExportIncludeHiddenSheets: React.Dispatch<React.SetStateAction<boolean>>;
  setExportIncludeSchemaTab: React.Dispatch<React.SetStateAction<boolean>>;
}

export function ExportDialog({ activeSheet, exportDialogRef, exportIncludeHidden, exportIncludeHiddenSheets, exportIncludeSchemaTab, exportUseMachineKeys, hiddenSheets, project, setExportIncludeHidden, setExportIncludeHiddenSheets, setExportIncludeSchemaTab }: ExportDialogProps) {
  return (
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
  );
}
