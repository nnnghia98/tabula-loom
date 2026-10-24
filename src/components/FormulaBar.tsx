import type { FieldConfig, CellValue } from '../types';
import type { ActiveCellInfo } from '../types';
import { IconPlus, IconSearch, IconSliders, IconWarning } from '../Icons';

interface FormulaBarProps {
  activeCellInfo: ActiveCellInfo | null;
  handleAddField: () => void;
  handleOpenAddRowModal: () => void;
  isInspectorOpen: boolean;
  searchQuery: string;
  setIsInspectorOpen: React.Dispatch<React.SetStateAction<boolean>>;
  setSearchQuery: React.Dispatch<React.SetStateAction<string>>;
  updateCellValue: (rowId: string, fieldId: string, rawValue: CellValue) => void;
  updateFieldConfig: (fieldId: string, patch: Partial<FieldConfig>) => void;
}

export function FormulaBar({ activeCellInfo, handleAddField, handleOpenAddRowModal, isInspectorOpen, searchQuery, setIsInspectorOpen, setSearchQuery, updateCellValue, updateFieldConfig }: FormulaBarProps) {
  return (
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
      <div className="formula-bar-actions">
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
  );
}
