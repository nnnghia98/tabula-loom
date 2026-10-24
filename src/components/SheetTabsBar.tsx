import type { SheetData, FieldConfig, RowRecord } from '../types';
import type { SheetErrors } from '../types';
import { IconPlus, IconEye } from '../Icons';

interface SheetTabsBarProps {
  activeSheet: SheetData;
  dismissedHiddenNotice: boolean;
  displayedSheets: SheetData[];
  filteredRows: RowRecord[];
  handleSelectSheet: (sheetId: string) => void;
  hiddenSheets: SheetData[];
  newSheetDialogRef: React.RefObject<HTMLDialogElement | null>;
  setShowHiddenSheets: React.Dispatch<React.SetStateAction<boolean>>;
  sheetErrorsMap: Record<string, SheetErrors>;
  showHiddenSheets: boolean;
  visibleFields: FieldConfig[];
}

export function SheetTabsBar({ activeSheet, dismissedHiddenNotice, displayedSheets, filteredRows, handleSelectSheet, hiddenSheets, newSheetDialogRef, setShowHiddenSheets, sheetErrorsMap, showHiddenSheets, visibleFields }: SheetTabsBarProps) {
  return (
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
      </div>
    </footer>
  );
}
